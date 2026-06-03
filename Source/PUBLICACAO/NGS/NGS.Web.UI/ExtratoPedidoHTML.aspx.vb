Imports System
Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ExtratoPedidoHTML
    Inherits BasePage

#Region "Estruturas"

    Private Structure structTotaisQuantidade
        Public QuantidadeContratada As Decimal
        Public QuantidadeFixada As Decimal
        Public QuantidadeEntregue As Decimal
        Public QuantidadeEntregueFisico As Decimal
        Public QuantidadeDevolvida As Decimal
        Public QuantidadeDevolvidaFisico As Decimal
        Public QuantidadeCessionario As Decimal
        Public QuantidadeCedente As Decimal
        Public ValorFixado As Decimal
        Public ValorPago As Decimal
        Public ValorNotaFiscal As Decimal
        Public tClasse As eClassesOperacoes
    End Structure

#End Region

#Region "Variáveis locais"

    Private objPedidos As New [Lib].Negocio.Pedidos()
    Private objQtdePedido As New structTotaisQuantidade()
    Private dateEmissao As DateTime = DateTime.Now
    Private strEmpresaQry, strClienteQry, strPedido, strGrupoProduto, strProduto, strNomeProduto, strDesPrd, strDataFim, strSafra As String
    Private blnEntrada As Boolean = False, blnSaida As Boolean = False
    Private intPagina As Integer = 0
    Private strNomePainel As String
    Private intLinhas As Integer = 0
    Private ProdutoAgrupado As Boolean = False
    Private SituacaoPedido As String

#End Region

#Region "Eventos"

    Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
        intPagina = 0

        If Request.QueryString("empresa") Is Nothing Then Response.Redirect("ExtratoDePedido.aspx")
        strEmpresaQry = Request.QueryString("empresa")
        'strClienteQry = Request.QueryString("cliente")
        If Not Request.QueryString("cliente") Is Nothing Then strClienteQry = Request.QueryString("cliente") Else strClienteQry = ""
        If Not Request.QueryString("pedido") Is Nothing Then strPedido = Request.QueryString("pedido") Else strPedido = ""
        If Not Request.QueryString("grupoproduto") Is Nothing Then strGrupoProduto = Request.QueryString("grupoproduto") Else strGrupoProduto = ""
        If Not Request.QueryString("produto") Is Nothing Then strProduto = Request.QueryString("produto") Else strProduto = ""
        If Not Request.QueryString("nomeproduto") Is Nothing Then strNomeProduto = Request.QueryString("nomeproduto") Else strNomeProduto = ""
        strDesPrd = Request.QueryString("desprd")
        If Not Request.QueryString("safra") Is Nothing Then strSafra = Request.QueryString("safra") Else strSafra = ""

        strDataFim = Request.QueryString("fim")

        Select Case Request.QueryString("es")
            Case "E" : blnEntrada = True
            Case "S" : blnSaida = True
            Case "ES"
                blnEntrada = True
                blnSaida = True
        End Select

        SituacaoPedido = Request.QueryString("tipo")

        Dim strHTML As String = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" & vbCrLf & _
                                "<html xmlns=""http://www.w3.org/1999/xhtml"">" & vbCrLf & _
                                "   <head>" & vbCrLf & _
                                "       <title>Extrato de Pedidos</title>" & vbCrLf & _
                                "       <style type=""text/css"">" & vbCrLf & _
                                "           body" & vbCrLf & _
                                "           {" & vbCrLf & _
                                "               font-size: 7pt;" & vbCrLf & _
                                "           }" & vbCrLf & _
                                "           .titulos" & vbCrLf & _
                                "           {" & vbCrLf & _
                                "               width: 100%;" & vbCrLf & _
                                "               border-top: solid 1px #000000;" & vbCrLf & _
                                "               border-bottom: solid 1px #000000;" & vbCrLf & _
                                "           }" & vbCrLf & _
                                "           .titulosbaixo" & vbCrLf & _
                                "           {" & vbCrLf & _
                                "               width: 100%;" & vbCrLf & _
                                "               border-bottom: solid 1px #000000;" & vbCrLf & _
                                "           }" & vbCrLf & _
                                "           .linhastitulobaixo" & vbCrLf & _
                                "           {" & vbCrLf & _
                                "               border-bottom: solid 1px #000000;" & vbCrLf & _
                                "           }" & vbCrLf & _
                                "           .linhastitulocima" & vbCrLf & _
                                "           {" & vbCrLf & _
                                "               border-top: solid 1px #000000;" & vbCrLf & _
                                "           }" & vbCrLf & _
                                "           .quebrapagina" & vbCrLf & _
                                "           {" & vbCrLf & _
                                "               page-break-after: auto;" & vbCrLf & _
                                "           }" & vbCrLf & _
                                "       </style>" & vbCrLf & _
                                "   </head>" & vbCrLf & _
                                "   <body>"

        writer.Write(strHTML)

        ImprimirPedidos(writer)

        strHTML = "   </body>" & _
                  "</html>"

        writer.Write(strHTML)
    End Sub

#End Region

#Region "Procedimentos"

#Region "Cabeçalho"

    Private Sub ImprimirCabecalho(ByRef writer As System.Web.UI.HtmlTextWriter, ByVal Pedido As [Lib].Negocio.Pedido)
        If dateEmissao = #12:00:00 AM# Then dateEmissao = DateTime.Now

        intPagina += 1
        intLinhas = 0

        Dim strCodigo As String() = strEmpresaQry.Split("-")

        Dim strSQL As String = "SELECT Nome, Cidade, Estado, Endereco, Inscricao " & _
                               "FROM Clientes " & _
                               "WHERE Cliente_Id = '" & strCodigo(0) & "' " & _
                               "AND Endereco_Id = " & strCodigo(1)

        Dim dsEmpresa As DataSet = Banco.ConsultaDataSet(strSQL, "Clientes")
        Dim drEmpresa As DataRow = dsEmpresa.Tables(0).Rows(0)

        Dim strEmpresa As String() = New String() {strCodigo(0), strCodigo(1), drEmpresa("Nome"), drEmpresa("Cidade"), _
                                                   drEmpresa("Estado"), drEmpresa("Endereco"), drEmpresa("Inscricao")}

        strSQL = "SELECT Nome, Cidade, Estado, Endereco, Complemento, Inscricao " & _
                 "FROM Clientes " & _
                 "WHERE Cliente_Id = '" & Pedido.CodigoCliente & "' " & _
                 "AND Endereco_Id = " & Pedido.EnderecoCliente

        Dim dsClientes As DataSet = Banco.ConsultaDataSet(strSQL, "Clientes")
        Dim drCliente As DataRow = dsClientes.Tables(0).Rows(0)
        Dim strCliente As String() = New String() {Pedido.CodigoCliente, Pedido.EnderecoCliente, drCliente("Nome"), drCliente("Cidade"), _
                                                   drCliente("Estado"), drCliente("Endereco"), drCliente("Complemento"), drCliente("Inscricao")}

        Dim strHTML As String = "<table width=""100%"">" & _
                                "<tr>" & _
                                "<td><b>" & strEmpresa(2) & "</b></td>" & _
                                "<td><b>P&aacute;gina</b></td>" & _
                                "<td>" & intPagina.ToString() & "</td>" & _
                                "</tr>" & _
                                "<tr>" & _
                                "<td>CNPJ: " & String.Format("{0:00\.000\.000\./0000\-00}", strEmpresa(0)) & " - Cidade: " & strEmpresa(3) & "-" & strEmpresa(4) & "</td>" & _
                                "<td><b>Data:</b></td>" & _
                                "<td>" & dateEmissao.ToString("dd/MM/yyyy") & "</td>" & _
                                "</tr>" & _
                                "<tr>" & _
                                "<td colspan=""3""><b>" & Funcoes.FormatarCpfCnpj(strCliente(0)) & " - " & strCliente(2) & " - " & strCliente(3) & "/" & strCliente(4) & "</b></td></tr>"

        If strCliente(6).ToString.Length > 0 Or strCliente(7).ToString.Length > 0 Then
            strHTML &= "<tr><td colspan=""3""><b>"

            If strCliente(7).ToString.Length > 0 Then
                strHTML &= "INSCR: " & strCliente(7) & ""
            End If
            If strCliente(6).ToString.Length > 0 Then
                strHTML &= " - " & strCliente(6) & "  "
            End If
            strHTML &= "</b></td></tr>"
        End If
        strHTML &= "</table>"

        writer.Write(strHTML)
    End Sub

#End Region

#Region "Pedidos"

    Private Function getTipoOperacao(ByVal objPedido As Pedido) As String
        Dim tipoOperacao As String = ""
        If objPedido IsNot Nothing Then
            If objPedido.Troca OrElse objPedido.Antecipada Then
                tipoOperacao = " - Tipo de Operação: "
                If objPedido.Troca Then
                    tipoOperacao &= "Troca "
                End If

                If objPedido.Antecipada Then
                    tipoOperacao &= "Antecipada"
                End If
            End If
        End If

        Return tipoOperacao
    End Function

    Private Sub ImprimirPedidos(ByRef writer As HtmlTextWriter)
        'Dim intLinha As Integer = 1
        intLinhas = 1

        SelecionarPedidos()

        For Each objPedido As [Lib].Negocio.Pedido In objPedidos
            Dim blnImprimir As Boolean = True

            objQtdePedido.QuantidadeContratada = 0
            objQtdePedido.QuantidadeFixada = 0
            objQtdePedido.QuantidadeEntregue = 0
            objQtdePedido.QuantidadeEntregueFisico = 0
            objQtdePedido.QuantidadeDevolvida = 0
            objQtdePedido.QuantidadeDevolvidaFisico = 0
            objQtdePedido.QuantidadeCessionario = 0
            objQtdePedido.QuantidadeCedente = 0
            objQtdePedido.ValorFixado = 0
            objQtdePedido.ValorPago = 0
            objQtdePedido.ValorNotaFiscal = 0

            Dim dsNotasFiscais As DataSet = Extrato.getDataSetNotasFiscais(objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa, objPedido.Codigo, "", 0, blnEntrada, blnSaida, strDataFim, SituacaoPedido)
            Dim dsResumoNotasFiscaisPrd As DataSet = Extrato.getDataSetResumoNotasFiscaisPrd(objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa, objPedido.Codigo, "", 0, blnEntrada, blnSaida, strDataFim, SituacaoPedido)
            Dim dsResumoNotasFiscais As DataSet = Extrato.getDataSetResumoNotasFiscais(objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa, objPedido.Codigo, "", 0, blnEntrada, blnSaida, strDataFim, SituacaoPedido)
            ''********************************************************************************************************************
            ''********************************************************************************************************************
            ''MARCELO, COLOQUEI ESSA ROTINA NOVAMENTE PORQUE NÃO ESTAVA LISTANDO OS ITENS ESTORNO/COMPLEMENTO DO(S) CONTRATO(S)
            ''VEJA O EMAIL QUE TE MANDEI DIA 03/05, ASSIM Q AJUSTAR PODE TIRAR - FURLAN - 04/05/2010
            objQtdePedido.tClasse = objPedido.SubOperacao.Classe
            Dim dsContratos As DataSet = Extrato.getDataSetContratos(objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa, objPedido.Codigo, "", 0, SituacaoPedido)

            ''********************************************************************************************************************
            ''********************************************************************************************************************
            Dim dsFinanceiro As DataSet = Extrato.getDataSetFinanceiro(objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa, objPedido.Codigo, "", 0, False, SituacaoPedido)

            Dim dsRazao As DataSet = SelecionarRazao(objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa, objPedido.Codigo)

            Dim dsAdiantamentos As DataSet = Extrato.getDataSetAdiantamentos(objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa, objPedido.Codigo, "", 0, SituacaoPedido)

            Dim dsProcuracoesCedidas As DataSet = Extrato.getDataSetProcuracoes(True, objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa, objPedido.Codigo, "", 0)
            Dim dsProcuracoesCessionarias As DataSet = Extrato.getDataSetProcuracoes(False, objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa, objPedido.Codigo, "", 0)

            'blnImprimir = (dsNotasFiscais.Tables(0).Rows.Count > 0)
            'If Not blnImprimir Then blnImprimir = objPedido.TemFixacoes
            'If Not blnImprimir Then blnImprimir = (dsFinanceiro.Tables(0).Rows.Count > 0)

            If blnImprimir Then
                ImprimirCabecalho(writer, objPedido)
                ImprimirCabecalhoPedido(writer, objPedido.Codigo, objPedido.CodigoOperacao, _
                                                objPedido.CodigoSubOperacao, objPedido.SubOperacao.Descricao, _
                                                objPedido.SubOperacao.EntradaSaida.ToString().Substring(0, 1), objPedido.CodigoSafra, IIf(objPedido.FreteCIFFOB = eTiposFrete.CIF, "CIF", "FOB"), _
                                                getTipoOperacao(objPedido), objPedido.Contrato)
            End If

            If dsNotasFiscais.Tables(0).Rows.Count > 0 Then ImprimirNotasFiscais(writer, dsNotasFiscais, dsResumoNotasFiscaisPrd, dsResumoNotasFiscais, objPedido)
            ''********************************************************************************************************************
            ''********************************************************************************************************************
            ''MARCELO, COLOQUEI ESSA ROTINA NOVAMENTE PORQUE NÃO ESTAVA LISTANDO OS ITENS ESTORNO/COMPLEMENTO DO(S) CONTRATO(S)
            ''VEJA O EMAIL QUE TE MANDEI DIA 03/05, ASSIM Q AJUSTAR PODE TIRAR - FURLAN - 04/05/2010
            If dsContratos.Tables(0).Rows.Count > 0 Then ImprimirContratos(writer, dsContratos, objPedido)
            'ImprimirContratos(writer, objPedido)
            ''********************************************************************************************************************
            ''********************************************************************************************************************
            If objPedido.TemFixacoes Then ImprimirFixacoes(writer, objPedido.Itens, objPedido)
            If dsFinanceiro.Tables(0).Rows.Count > 0 Then ImprimirFinanceiro(writer, dsFinanceiro, objPedido)

            If dsRazao.Tables(0).Rows.Count > 0 Then ImprimirRazao(writer, dsRazao, objPedido)

            If Not FinanceiroNovo AndAlso dsAdiantamentos.Tables(0).Rows.Count > 0 Then ImprimirAdiantamentos(writer, dsAdiantamentos, objPedido)

            If dsProcuracoesCedidas.Tables(0).Rows.Count > 0 Then ImprimirProcuracoes(writer, dsProcuracoesCedidas, True, objPedido)
            If dsProcuracoesCessionarias.Tables(0).Rows.Count > 0 Then ImprimirProcuracoes(writer, dsProcuracoesCessionarias, False, objPedido)

            ImprimirTotais(writer, objPedido)

            If blnImprimir Then
                writer.Write("</td></tr></table>")
                'If intLinhas <> objPedidos.Count Then writer.Write("<br class=""quebrapagina"" />")
                If objPedidos.Count > 1 Then writer.Write("<br class=""quebrapagina"" />")
                'intLinhas = intLinhas + 1
                intLinhas = 1
            End If
        Next
    End Sub

    Private Sub ImprimirCabecalhoPedido(ByRef writer As HtmlTextWriter, ByVal Pedido As Integer, ByVal Operacao As Integer, _
                                        ByVal SubOperacao As Integer, ByVal OperacaoDesc As String, ByVal EntradaSaida As String, ByVal Safra As String, ByVal CIFFOB As String, ByVal TipoOperacao As String, ByVal Contrato As String)

        Dim strHTML As String = "<table width=""100%"" style=""border: solid 1px #000000"">" & _
                                "<tr>" & _
                                "<td>" & _
                                "<b>PEDIDO: " & Pedido & "</b> - " & Operacao & "-" & SubOperacao & " - " & OperacaoDesc & " - (" & EntradaSaida & ")" & " - " & Safra & _
                                " - FRETE: " & CIFFOB & TipoOperacao & _
                                "</td>" & _
                                "</tr>"

        If Not String.IsNullOrWhiteSpace(Contrato) Then
            strHTML &= "<tr>" & _
                       "<td>" & _
                       "<b>CONTRATO: " & Contrato & "</b>" & _
                       "</td>" & _
                       "</tr>"
        End If

        strHTML &= "<tr>" & _
                   "<td>"

        writer.Write(strHTML)
    End Sub

    'Private Sub ImprimirCabecalhoContinuacao(ByRef writer As HtmlTextWriter)
    '    Dim strHTML As String = "<table width=""100%"" style=""border: solid 1px #000000"">" & _
    '                            "<tr>" & _
    '                            "<td>"

    '    writer.Write(strHTML)
    'End Sub


#End Region

#Region "Notas Fiscais"

    Private Sub ImprimirNotasFiscais(ByVal writer As HtmlTextWriter, ByVal Dados As DataSet, ByVal ResumoDadosPrd As DataSet, ByVal ResumoDados As DataSet, ByVal Pedido As [Lib].Negocio.Pedido)
        Dim dsNotasFiscais As DataSet = Dados
        Dim NotaPedido As New [Lib].Negocio.Pedido(Pedido.CodigoEmpresa, Pedido.EnderecoEmpresa, Pedido.Codigo)
        Dim intSinal As Integer


        If (dsNotasFiscais.Tables(0).Rows.Count > 0) Then
            Dim blnMudar As Boolean = False
            Dim dblSomas As Double() = {0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0}
            Dim strHTML As String

            'Dim objProduto As New [Lib].Negocio.Produto(dsNotasFiscais.Tables(0).Rows(0).Item("Produto"))
            'If objProduto.Agrupar = "S" Then ProdutoAgrupado = True
            If Not IsDBNull(dsNotasFiscais.Tables(0).Compute("SUM(PesoFiscal)", "AGRUPAR = 'S'")) Then ProdutoAgrupado = True

            Dim temRemessa As Boolean = False
            If Not IsDBNull(dsNotasFiscais.Tables(0).Compute("SUM(PesoFiscal)", "CLASSE = 'REMESSAS'")) Then temRemessa = True

            ImprimirCabecalhoNotasFiscais(writer)

            For Each drNotaFiscal As DataRow In dsNotasFiscais.Tables(0).Rows
                If drNotaFiscal("Classe") = "AFIXAR" And drNotaFiscal("Devolucao") = "S" Then
                    intSinal = -1
                ElseIf drNotaFiscal("Classe") = "AFIXAR" Then
                    intSinal = 1
                ElseIf NotaPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada And drNotaFiscal("EntradaSaida").ToString() = "S" Then
                    intSinal = -1
                ElseIf NotaPedido.SubOperacao.EntradaSaida = eEntradaSaida.Saida And drNotaFiscal("EntradaSaida").ToString() = "E" Then
                    intSinal = -1
                Else
                    intSinal = 1
                End If

                If blnMudar Then
                    writer.Write("</table>")
                    writer.Write("</td></tr></table>")
                    writer.Write("<br class=""quebrapagina"" />")
                    ImprimirCabecalho(writer, Pedido)
                    ImprimirCabecalhoPedido(writer, Pedido.Codigo, Pedido.CodigoOperacao, _
                                                    Pedido.CodigoSubOperacao, Pedido.SubOperacao.Descricao, _
                                                    Pedido.SubOperacao.EntradaSaida.ToString().Substring(0, 1), Pedido.CodigoSafra, Pedido.FreteCIFFOB, getTipoOperacao(Pedido), Pedido.Contrato)
                    ImprimirCabecalhoNotasFiscais(writer)
                    blnMudar = False
                End If

                strHTML = "<tr>" & _
                          "<td align=""left"">" & _
                          Convert.ToDateTime(drNotaFiscal("Movimento")).ToString("dd/MM/yyy") & _
                          "</td>"

                If strDesPrd = "S" Then
                    strHTML &= "<td align=""left"">" & _
                              drNotaFiscal("Produto").ToString() & _
                              "-" & _
                              drNotaFiscal("NomeProduto").ToString() & _
                              "</td>"
                Else
                    strHTML &= "<td align=""left"">" & _
                              drNotaFiscal("Produto").ToString() & _
                              "</td>"
                End If

                If Not drNotaFiscal.IsNull("Operacao") Then
                    strHTML &= "<td align=""center"">" & _
                                Convert.ToInt32(drNotaFiscal("Operacao")).ToString("00") & "-" & Convert.ToInt32(drNotaFiscal("SubOperacao")).ToString("00") & _
                                "</td>"
                Else : strHTML &= "<td align=""center"">&nbsp;</td>"
                End If

                If Not drNotaFiscal.IsNull("CFOP") Then
                    strHTML &= "<td align=""center"">" & _
                                Convert.ToInt32(drNotaFiscal("CFOP")).ToString("0000") & _
                                "</td>"
                Else : strHTML &= "<td align=""center"">&nbsp;</td>"
                End If

                If Not drNotaFiscal.IsNull("Finalidade") Then
                    strHTML &= "<td align=""center"">" & _
                                Convert.ToInt32(drNotaFiscal("Finalidade")).ToString("000") & _
                                "</td>"
                Else : strHTML &= "<td align=""center"">&nbsp;</td>"
                End If

                strHTML &= "<td align=""right"">" & _
                          drNotaFiscal("Reduzido").ToString() & _
                          "</td>"

                If Not drNotaFiscal.IsNull("Nota") Then
                    strHTML &= "<td align=""center"">" & _
                               drNotaFiscal("EntradaSaida").ToString() & "-" & drNotaFiscal("Serie").ToString() & _
                               "</td>" & _
                               "<td align=""right"">" & _
                               drNotaFiscal("Nota").ToString() & _
                               "</td>"
                Else
                    strHTML &= "<td align=""center"">&nbsp;</td>" & _
                               "<td align=""right"">&nbsp;</td>"
                End If

                strHTML &= "<td align=""left"">" & _
                          Trim(drNotaFiscal("Placa")) & _
                          "</td>"

                writer.Write(strHTML)

                If ProdutoAgrupado = False Then
                    strHTML = "<td align=""right"">" & _
                              drNotaFiscal("NumeroTicket").ToString() & _
                              "</td>" & _
                              "<td align=""right"">"

                    writer.Write(strHTML)

                    If Not drNotaFiscal.IsNull("PesoBalanca") Then
                        writer.Write(Convert.ToInt32(drNotaFiscal("PesoBalanca")).ToString("N0"))
                        dblSomas(0) += Convert.ToDouble(drNotaFiscal("PesoBalanca")) * intSinal
                    Else : writer.Write("&nbsp;")
                    End If

                    writer.Write("</td><td align=""right"">")

                    If Not drNotaFiscal.IsNull("Umidade") Then
                        writer.Write(Convert.ToInt32(drNotaFiscal("Umidade")).ToString("N0"))
                        dblSomas(1) += Convert.ToDouble(drNotaFiscal("Umidade")) * intSinal
                    Else : writer.Write("&nbsp;")
                    End If

                    writer.Write("</td><td align=""right"">")

                    If Not drNotaFiscal.IsNull("Impureza") Then
                        writer.Write(Convert.ToInt32(drNotaFiscal("Impureza")).ToString("N0"))
                        dblSomas(2) += Convert.ToDouble(drNotaFiscal("Impureza")) * intSinal
                    Else : writer.Write("&nbsp;")
                    End If

                    writer.Write("</td><td align=""right"">")

                    If Not drNotaFiscal.IsNull("Avariados") Then
                        writer.Write(Convert.ToInt32(drNotaFiscal("Avariados")).ToString("N0"))
                        dblSomas(3) += Convert.ToDouble(drNotaFiscal("Avariados")) * intSinal
                    Else : writer.Write("&nbsp;")
                    End If

                    writer.Write("</td><td align=""right"">")

                    If Not drNotaFiscal.IsNull("PH") Then
                        writer.Write(Convert.ToInt32(drNotaFiscal("PH")).ToString("N0"))
                        dblSomas(4) += Convert.ToDouble(drNotaFiscal("PH")) * intSinal
                    Else : writer.Write("&nbsp;")
                    End If

                    writer.Write("</td><td align=""right"">")

                    If Not drNotaFiscal.IsNull("GMO") Then writer.Write(Convert.ToInt32(drNotaFiscal("GMO")).ToString("N0")) Else writer.Write("&nbsp;")

                    writer.Write("</td>")

                    writer.Write("<td align=""right"">")

                    If Not drNotaFiscal.IsNull("PesoLiquido") Then
                        writer.Write(Convert.ToInt32(drNotaFiscal("PesoLiquido")).ToString("N0"))
                        dblSomas(5) += Convert.ToDouble(drNotaFiscal("PesoLiquido")) * intSinal
                    Else : writer.Write("&nbsp;")
                    End If

                    writer.Write("</td><td align=""right"">")

                    If Not drNotaFiscal.IsNull("PesoFiscal") Then
                        writer.Write(Convert.ToInt32(drNotaFiscal("PesoFiscal")).ToString("N0"))

                        If drNotaFiscal("Classe") = "GLOBAL" Then
                            If Not temRemessa Then dblSomas(6) += Convert.ToDouble(drNotaFiscal("PesoFiscal")) * intSinal
                        Else
                            dblSomas(6) += Convert.ToDouble(drNotaFiscal("PesoFiscal")) * intSinal

                            If intSinal = 1 Then
                                If (Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada AndAlso Pedido.FreteCIFFOB = eTiposFrete.CIF) OrElse _
                                   (Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso Pedido.FreteCIFFOB = eTiposFrete.FOB) Then
                                    If drNotaFiscal("PesoLiquido") = 0 Then
                                        objQtdePedido.QuantidadeEntregue += Convert.ToDouble(drNotaFiscal("PesoFiscal"))
                                        'objQtdePedido.QuantidadeEntregueFisico += Convert.ToDouble(drNotaFiscal("PesoFiscal"))
                                    Else
                                        objQtdePedido.QuantidadeEntregueFisico += Convert.ToDouble(drNotaFiscal("PesoLiquido"))
                                        objQtdePedido.QuantidadeEntregue += Convert.ToDouble(drNotaFiscal("PesoFiscal"))
                                    End If
                                Else
                                    objQtdePedido.QuantidadeEntregue += Convert.ToDouble(drNotaFiscal("PesoFiscal"))
                                    objQtdePedido.QuantidadeEntregueFisico += Convert.ToDouble(drNotaFiscal("PesoFiscal"))
                                End If
                            Else
                                If drNotaFiscal("PesoLiquido") = 0 Then
                                    objQtdePedido.QuantidadeDevolvida += Convert.ToDouble(drNotaFiscal("PesoFiscal"))
                                    'objQtdePedido.QuantidadeDevolvidaFisico += Convert.ToDouble(drNotaFiscal("PesoLiquido"))
                                Else
                                    objQtdePedido.QuantidadeDevolvida += Convert.ToDouble(drNotaFiscal("PesoFiscal"))
                                    objQtdePedido.QuantidadeDevolvidaFisico += Convert.ToDouble(drNotaFiscal("PesoLiquido"))
                                End If
                            End If
                        End If
                    Else : writer.Write("&nbsp;")
                    End If
                Else
                    writer.Write("<td align=""right"">")

                    If Not drNotaFiscal.IsNull("PesoLiquido") Then
                        writer.Write(Convert.ToDouble(drNotaFiscal("PesoLiquido")).ToString("N4"))
                        dblSomas(5) += Convert.ToDouble(drNotaFiscal("PesoLiquido")) * intSinal
                    Else : writer.Write("&nbsp;")
                    End If

                    writer.Write("</td><td align=""right"">")

                    If Not drNotaFiscal.IsNull("PesoFiscal") Then
                        writer.Write(Convert.ToDouble(drNotaFiscal("PesoFiscal")).ToString("N4"))

                        If drNotaFiscal("Classe") = "GLOBAL" Then
                            If Not temRemessa Then dblSomas(6) += Convert.ToDouble(drNotaFiscal("PesoFiscal")) * intSinal
                        Else
                            dblSomas(6) += Convert.ToDouble(drNotaFiscal("PesoFiscal")) * intSinal

                            If intSinal = 1 Then
                                If (Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada AndAlso Pedido.FreteCIFFOB = eTiposFrete.CIF) OrElse _
                                   (Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso Pedido.FreteCIFFOB = eTiposFrete.FOB) Then
                                    If drNotaFiscal("PesoLiquido") = 0 Then
                                        objQtdePedido.QuantidadeEntregue += Convert.ToDouble(drNotaFiscal("PesoFiscal"))
                                        'objQtdePedido.QuantidadeEntregueFisico += Convert.ToDouble(drNotaFiscal("PesoFiscal"))
                                    Else
                                        objQtdePedido.QuantidadeEntregueFisico += Convert.ToDouble(drNotaFiscal("PesoLiquido"))
                                        objQtdePedido.QuantidadeEntregue += Convert.ToDouble(drNotaFiscal("PesoFiscal"))
                                    End If
                                Else
                                    objQtdePedido.QuantidadeEntregue += Convert.ToDouble(drNotaFiscal("PesoFiscal"))
                                End If
                            Else
                                If drNotaFiscal("PesoLiquido") = 0 Then
                                    objQtdePedido.QuantidadeDevolvida += Convert.ToDouble(drNotaFiscal("PesoFiscal"))
                                    'objQtdePedido.QuantidadeDevolvidaFisico += Convert.ToDouble(drNotaFiscal("PesoLiquido"))
                                Else
                                    objQtdePedido.QuantidadeDevolvida += Convert.ToDouble(drNotaFiscal("PesoFiscal"))
                                    objQtdePedido.QuantidadeDevolvidaFisico += Convert.ToDouble(drNotaFiscal("PesoLiquido"))
                                End If
                            End If
                        End If
                    Else : writer.Write("&nbsp;")
                    End If
                End If

                writer.Write("</td><td align=""right"">")

                If Not drNotaFiscal.IsNull("Unitario") Then writer.Write(Convert.ToDouble(drNotaFiscal("Unitario")).ToString("N10")) Else writer.Write("&nbsp;")

                writer.Write("</td><td align=""right"">")

                'If NotaPedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR Then
                '    intSinal = 1
                'End If

                If Not drNotaFiscal.IsNull("Valor") Then
                    writer.Write(Convert.ToDouble(drNotaFiscal("Valor")).ToString("N2"))

                    If drNotaFiscal("Classe") = "GLOBAL" Then
                        If Not temRemessa Then dblSomas(7) += Convert.ToDouble(drNotaFiscal("Valor")) * intSinal
                    Else
                        dblSomas(7) += Convert.ToDouble(drNotaFiscal("Valor")) * intSinal
                    End If
                Else : writer.Write("&nbsp;")
                End If

                writer.Write("</td><td align=""right"">")

                If Not drNotaFiscal.IsNull("Liquido") Then
                    writer.Write(Convert.ToDouble(drNotaFiscal("Liquido")).ToString("N2"))

                    If drNotaFiscal("Classe") = "GLOBAL" Then
                        If Not temRemessa Then dblSomas(8) += Convert.ToDouble(drNotaFiscal("Liquido")) * intSinal
                    Else
                        dblSomas(8) += Convert.ToDouble(drNotaFiscal("Liquido")) * intSinal
                    End If

                    If NotaPedido.SubOperacao.Classe = eClassesOperacoes.GLOBAL Then
                        If drNotaFiscal("Classe") = "GLOBAL" Then objQtdePedido.ValorNotaFiscal += Convert.ToDouble(drNotaFiscal("Liquido")) * intSinal
                    Else
                        objQtdePedido.ValorNotaFiscal = dblSomas(8)
                    End If
                Else : writer.Write("&nbsp;")
                End If

                writer.Write("</td><td align=""right""><a id=""imgEncargos"" alt="""" height=""15"" width=""15"" title=""Encargos da Nota Fiscal"" style=""cursor: pointer"" onclick=""window.open('EncargosDaNota.aspx?tipo=NOTA&emp=" &
                drNotaFiscal("Empresa").ToString() & "&endemp=" & drNotaFiscal("EndEmpresa").ToString() &
                "&cli=" & drNotaFiscal("Cliente").ToString() &
                "&endcli=" & drNotaFiscal("EndCliente").ToString() &
                "&entsai=" & drNotaFiscal("EntradaSaida").ToString() &
                "&serie=" & drNotaFiscal("Serie").ToString() &
                "&nota=" & drNotaFiscal("Nota").ToString() &
                "&produto=" & drNotaFiscal("Produto").ToString() & "', '', 'resizable=no, menubar=no, scrollbars=No, width=500, height=280, top=100, left=200');""><i class='fa fa-arrow-right seta'></i></a></td></tr>")

                intLinhas += 1

                If strDesPrd = "S" Then
                    If intLinhas = 38 Then
                        blnMudar = True
                        intLinhas = 0
                    End If
                Else
                    If intLinhas = 45 Then
                        blnMudar = True
                        intLinhas = 0
                    End If
                End If
            Next

            ImprimirTotaisNotasFiscais(writer, dblSomas, Pedido)

            If ResumoDados.Tables(0).Rows.Count > 1 Then
                ImprimirResumoNotasFiscais(writer, ResumoDados, Pedido)
            End If

            If ResumoDadosPrd.Tables(0).Rows.Count > 1 Then
                ImprimirResumoNotasFiscaisPrd(writer, ResumoDadosPrd, Pedido)
            End If

            writer.Write("</table>")
        End If
    End Sub

    Private Sub ImprimirCabecalhoNotasFiscais(ByRef writer As System.Web.UI.HtmlTextWriter)
        Dim strHTML As String = "<table class=""titulos"">" & _
                                "<tr>" & _
                                "<td>NOTAS FISCAIS</td>" & _
                                "</tr>" & _
                                "</table>"

        writer.Write(strHTML)

        If ProdutoAgrupado = False Then
            strHTML = "<table class=""titulos"">" & _
                      "<tr>" & _
                      "<td class=""linhastitulobaixo"" rowspan=""2"">" & _
                      "Data<br />Movto" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" rowspan=""2"" align=""left"">" & _
                      "&nbsp;<br />Produto" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" rowspan=""2"" align=""center"">" & _
                      "Oper.<br />Coml." & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" rowspan=""2"" align=""center"">" & _
                      "&nbsp;<br />CFOP" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" rowspan=""2"" align=""center"">" & _
                      "Fin.<br />" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" rowspan=""2"" align=""right"">" & _
                      "Dep.<br />Efet." & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" rowspan=""2"" align=""center"">" & _
                      "E/S<br />Ser" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" rowspan=""2"" align=""right"">" & _
                      "Nota<br />Fiscal" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" rowspan=""2"" align=""left"">" & _
                      "Placa" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" rowspan=""2"" align=""right"">" & _
                      "N&uacute;mero<br />Ticket" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" rowspan=""2"" align=""right"">" & _
                      "Peso<br />Balan&ccedil;a" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" colspan=""5"" align=""center"">" & _
                      "DESCONTOS" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" rowspan=""2"" align=""right"">" & _
                      "Peso<br />L&iacute;quido" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" rowspan=""2"" align=""right"">" & _
                      "Peso<br />N.Fiscal" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" rowspan=""2"" align=""right"">" & _
                      "&nbsp;<br />Unit" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" rowspan=""2"" align=""right"">" & _
                      "Valor<br />Nota Fiscal" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" rowspan=""2"" align=""right"">" & _
                      "Liquido<br />Nota Fiscal" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" rowspan=""2"" align=""right"">&nbsp;</td>" & _
                      "</tr>"
        Else
            strHTML = "<table class=""titulos"">" & _
                      "<tr>" & _
                      "<td class=""linhastitulobaixo"">" & _
                      "Data<br />Movto" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" align=""left"">" & _
                      "&nbsp;<br />Produto" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" align=""center"">" & _
                      "Oper.<br />Coml." & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" align=""center"">" & _
                      "&nbsp;<br />CFOP" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" align=""center"">" & _
                      "Fin.<br />" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" align=""right"">" & _
                      "Dep.<br />Efet." & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" align=""center"">" & _
                      "E/S<br />Ser" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" align=""right"">" & _
                      "Nota<br />Fiscal" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" align=""left"">" & _
                      "Placa<br />" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" align=""right"">" & _
                      "Peso<br />L&iacute;quido" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" align=""right"">" & _
                      "Peso<br />N.Fiscal" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" align=""right"">" & _
                      "&nbsp;<br />Unit" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" align=""right"">" & _
                      "Valor<br />Nota Fiscal" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" align=""right"">" & _
                      "Liquido<br />Nota Fiscal" & _
                      "</td>" & _
                      "<td class=""linhastitulobaixo"" align=""right"">&nbsp;</td>" & _
                      "</tr>"
        End If

        writer.Write(strHTML)

        If ProdutoAgrupado = False Then
            strHTML = "<tr>" & _
                       "<td class=""linhastitulobaixo"" align=""center"">" & _
                       "Umidade" & _
                       "</td>" & _
                       "<td class=""linhastitulobaixo"" align=""center"">" & _
                       "Impur." & _
                       "</td>" & _
                       "<td class=""linhastitulobaixo"" align=""center"">" & _
                       "Avar." & _
                       "</td>" & _
                       "<td class=""linhastitulobaixo"" align=""center"">" & _
                       "P.H." & _
                       "</td>" & _
                       "<td class=""linhastitulobaixo"" align=""center"">" & _
                       "GMO" & _
                       "</td>" & _
                       "</tr>"
            writer.Write(strHTML)
        End If
    End Sub

    Private Sub ImprimirTotaisNotasFiscais(ByRef writer As System.Web.UI.HtmlTextWriter, ByVal Somas As Double(), ByVal Pedido As [Lib].Negocio.Pedido)
        Dim blnMudar As Boolean = False

        Dim strEmpresa As String() = strEmpresaQry.Split("-")
        Dim strCliente As String() = strClienteQry.Split(";")

        If strDesPrd = "S" Then
            If intLinhas = 38 Then
                blnMudar = True
                intLinhas = 0
            End If
        Else
            If intLinhas = 45 Then
                blnMudar = True
                intLinhas = 0
            End If
        End If

        If blnMudar Then
            writer.Write("</table>")
            writer.Write("</td></tr></table>")
            writer.Write("<br class=""quebrapagina"" />")
            ImprimirCabecalho(writer, Pedido)
            ImprimirCabecalhoPedido(writer, Pedido.Codigo, Pedido.CodigoOperacao, _
                                            Pedido.CodigoSubOperacao, Pedido.SubOperacao.Descricao, _
                                            Pedido.SubOperacao.EntradaSaida.ToString().Substring(0, 1), Pedido.CodigoSafra, Pedido.FreteCIFFOB, getTipoOperacao(Pedido), Pedido.Contrato)
            ImprimirCabecalhoNotasFiscais(writer)
            blnMudar = False
        End If

        Dim strHTML As String = ""

        If ProdutoAgrupado = False Then
            strHTML = "<tr>" &
                      "<td colspan=""10"" class=""linhastitulocima"">" &
                      "T O T A I S" &
                      "</td>" &
                      "<td class=""linhastitulocima"" align=""right"">" &
                      Convert.ToInt32(Somas(0)).ToString("N0") &
                      "</td>" &
                      "<td class=""linhastitulocima"" align=""right"">" &
                      Convert.ToInt32(Somas(1)).ToString("N0") &
                      "</td>" &
                      "<td class=""linhastitulocima"" align=""right"">" &
                      Convert.ToInt32(Somas(2)).ToString("N0") &
                      "</td>" &
                      "<td class=""linhastitulocima"" align=""right"">" &
                      Convert.ToInt32(Somas(3)).ToString("N0") &
                      "</td>" &
                      "<td class=""linhastitulocima"" align=""right"">" &
                      Convert.ToInt32(Somas(4)).ToString("N0") &
                      "</td>" &
                      "<td class=""linhastitulocima"" align=""right"">" &
                      "&nbsp;" &
                      "</td>" &
                      "<td class=""linhastitulocima"" align=""right"">" &
                      Convert.ToInt32(Somas(5)).ToString("N0") &
                      "</td>" &
                      "<td class=""linhastitulocima"" align=""right"">" &
                      Convert.ToInt32(Somas(6)).ToString("N0") &
                      "</td>" &
                      "<td class=""linhastitulocima"" align=""right"">" &
                      "&nbsp;" &
                      "</td>" &
                      "<td class=""linhastitulocima"" align=""right"">" &
                      Convert.ToDouble(Somas(7)).ToString("N2") &
                      "</td>" &
                      "<td class=""linhastitulocima"" align=""right"">" &
                      Convert.ToDouble(Somas(8)).ToString("N2") &
                      "</td>" &
                      "<td class=""linhastitulocima"" align=""right""><a id=""imgEncargos"" alt="""" height=""15"" width=""15"" title=""Total dos Encargos da Nota Fiscal"" style=""cursor: pointer"" onclick=""window.open('EncargosDaNota.aspx?tipo=TNOTA&emp=" &
                      strEmpresa(0).ToString() & "&endemp=" & strEmpresa(1).ToString() &
                      "&cli=" & Pedido.CodigoCliente.ToString() &
                      "&endcli=" & Pedido.EnderecoCliente.ToString() &
                      "&pedido=" & Pedido.Codigo.ToString & "', '', 'resizable=no, menubar=no, scrollbars=No, width=500, height=280, top=100, left=200');""><i class='fa fa-arrow-right seta'></i> </a></td></tr>"
        Else
            strHTML = "<tr>" &
                      "<td colspan=""9"" class=""linhastitulocima"">" &
                      "T O T A I S" &
                      "</td>" &
                      "<td class=""linhastitulocima"" align=""right"">" &
                      Convert.ToDouble(Somas(5)).ToString("N4") &
                      "</td>" &
                      "<td class=""linhastitulocima"" align=""right"">" &
                      Convert.ToDouble(Somas(6)).ToString("N4") &
                      "</td>" &
                      "<td class=""linhastitulocima"" align=""right"">" &
                      "&nbsp;" &
                      "</td>" &
                      "<td class=""linhastitulocima"" align=""right"">" &
                      Convert.ToDouble(Somas(7)).ToString("N2") &
                      "</td>" &
                      "<td class=""linhastitulocima"" align=""right"">" &
                      Convert.ToDouble(Somas(8)).ToString("N2") &
                      "</td>" &
                      "<td class=""linhastitulocima"" align=""right""><a id=""imgEncargos"" alt="""" height=""15"" width=""15"" title=""Total dos Encargos da Nota Fiscal"" style=""cursor: pointer"" onclick=""window.open('EncargosDaNota.aspx?tipo=TNOTA&emp=" &
                      strEmpresa(0).ToString() & "&endemp=" & strEmpresa(1).ToString() &
                      "&cli=" & Pedido.CodigoCliente.ToString() &
                      "&endcli=" & Pedido.EnderecoCliente.ToString() &
                      "&pedido=" & Pedido.Codigo.ToString & "', '', 'resizable=no, menubar=no, scrollbars=No, width=500, height=280, top=100, left=200');""><i class='fa fa-arrow-right seta'></i></a></td></tr>"
        End If

        writer.Write(strHTML)

        intLinhas += 2

        If ProdutoAgrupado = False And ((Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada AndAlso Pedido.FreteCIFFOB = eTiposFrete.CIF) OrElse _
           (Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso Pedido.FreteCIFFOB = eTiposFrete.FOB)) Then
            If Somas(5) < Somas(6) Then
                strHTML = "<tr>" & _
                          "<td align=""center"" colspan=""22"" class=""linhastitulocima"">" & _
                          "* * * * * * * * * *  ATENÇÃO, FALTA EMITIR NOTA FISCAL DE DEVOLUÇÃO  * * * * * * * * * *" & _
                          "</td></tr>"
                writer.Write(strHTML)

                intLinhas += 3
            End If
        End If
    End Sub

    Private Sub ImprimirResumoNotasFiscaisPrd(ByRef writer As System.Web.UI.HtmlTextWriter, ByVal ResumoDadosPrd As DataSet, ByVal Pedido As [Lib].Negocio.Pedido)
        Dim blnMudar As Boolean = False

        If strDesPrd = "S" Then
            If intLinhas = 38 Then
                blnMudar = True
                intLinhas = 0
            End If
        Else
            If intLinhas = 45 Then
                blnMudar = True
                intLinhas = 0
            End If
        End If

        Dim strEmpresa As String() = strEmpresaQry.Split("-")
        Dim strCliente As String() = strClienteQry.Split(";")

        If (ResumoDadosPrd.Tables(0).Rows.Count > 0) Then
            Dim strHTML As String = ""
            Dim primeiro As Boolean = True

            For Each drResumoNotaFiscalPrd As DataRow In ResumoDadosPrd.Tables(0).Rows
                If blnMudar Then
                    writer.Write("</table>")
                    writer.Write("</td></tr></table>")
                    writer.Write("<br class=""quebrapagina"" />")
                    ImprimirCabecalho(writer, Pedido)
                    ImprimirCabecalhoPedido(writer, Pedido.Codigo, Pedido.CodigoOperacao, _
                                                    Pedido.CodigoSubOperacao, Pedido.SubOperacao.Descricao, _
                                                    Pedido.SubOperacao.EntradaSaida.ToString().Substring(0, 1), Pedido.CodigoSafra, Pedido.FreteCIFFOB, getTipoOperacao(Pedido), Pedido.Contrato)
                    ImprimirCabecalhoNotasFiscais(writer)
                    blnMudar = False
                    primeiro = True
                End If

                If primeiro Then
                    If ProdutoAgrupado = False Then
                        strHTML = "<tr>" & _
                                    "<td class=""linhastitulocima"" colspan=""22"">" & _
                                    IIf(primeiro, "TOTAL POR PRODUTO", "&nbsp;") & _
                                    "</td></tr>"
                    Else
                        strHTML = "<tr>" & _
                                    "<td class=""linhastitulocima"" colspan=""15"">" & _
                                    IIf(primeiro, "TOTAL POR PRODUTO", "&nbsp;") & _
                                    "</td></tr>"
                    End If

                    writer.Write(strHTML)

                    intLinhas += 1
                End If

                strHTML = "<tr>" & _
                            "<td>&nbsp;</td>" & _
                            "<td align=""left"">" & _
                            drResumoNotaFiscalPrd("Produto") & "-" & drResumoNotaFiscalPrd("Nome") & _
                            "</td>" & _
                            "<td colspan=""7"">" & _
                            "&nbsp;" & _
                            "</td>" & _
                            "<td align=""right"">" & _
                            Convert.ToDouble(drResumoNotaFiscalPrd("PesoLiquido")).ToString("N4") & _
                            "</td>" & _
                            "<td align=""right"">" & _
                            Convert.ToDouble(drResumoNotaFiscalPrd("PesoFiscal")).ToString("N4") & _
                            "</td>" & _
                            "<td align=""right"">" & _
                            "&nbsp;" & _
                            "</td>" & _
                            "<td align=""right"">" & _
                            Convert.ToDouble(drResumoNotaFiscalPrd("Valor")).ToString("N2") & _
                            "</td>" & _
                            "<td align=""right"">" & _
                            Convert.ToDouble(drResumoNotaFiscalPrd("Liquido")).ToString("N2") & _
                            "</td>" & _
                            "<td align=""right"">" & _
                            "&nbsp;</td></tr>"

                writer.Write(strHTML)

                intLinhas += 1

                If strDesPrd = "S" Then
                    If intLinhas = 38 Then
                        blnMudar = True
                        intLinhas = 0
                    End If
                Else
                    If intLinhas = 45 Then
                        blnMudar = True
                        intLinhas = 0
                    End If
                End If

                primeiro = False
            Next
        End If
    End Sub

    Private Sub ImprimirResumoNotasFiscais(ByRef writer As System.Web.UI.HtmlTextWriter, ByVal ResumoDados As DataSet, ByVal Pedido As [Lib].Negocio.Pedido)
        Dim blnMudar As Boolean = False

        Dim strEmpresa As String() = strEmpresaQry.Split("-")
        Dim strCliente As String() = strClienteQry.Split(";")

        If strDesPrd = "S" Then
            If intLinhas = 38 Then
                blnMudar = True
                intLinhas = 0
            End If
        Else
            If intLinhas = 45 Then
                blnMudar = True
                intLinhas = 0
            End If
        End If

        If (ResumoDados.Tables(0).Rows.Count > 0) Then
            Dim strHTML As String = ""
            Dim primeiro As Boolean = True

            For Each drResumoNotaFiscal As DataRow In ResumoDados.Tables(0).Rows
                If blnMudar Then
                    writer.Write("</table>")
                    writer.Write("</td></tr></table>")
                    writer.Write("<br class=""quebrapagina"" />")
                    ImprimirCabecalho(writer, Pedido)
                    ImprimirCabecalhoPedido(writer, Pedido.Codigo, Pedido.CodigoOperacao, _
                                                    Pedido.CodigoSubOperacao, Pedido.SubOperacao.Descricao, _
                                                    Pedido.SubOperacao.EntradaSaida.ToString().Substring(0, 1), Pedido.CodigoSafra, Pedido.FreteCIFFOB, getTipoOperacao(Pedido), Pedido.Contrato)
                    ImprimirCabecalhoNotasFiscais(writer)
                    blnMudar = False
                    primeiro = True
                End If

                If primeiro Then
                    If ProdutoAgrupado = False Then
                        strHTML = "<tr>" & _
                                    "<td class=""linhastitulocima"" colspan=""22"">" & _
                                    IIf(primeiro, "TOTAL POR OPERAÇÃO", "&nbsp;") & _
                                    "</td></tr>"
                    Else
                        strHTML = "<tr>" & _
                                    "<td class=""linhastitulocima"" colspan=""15"">" & _
                                    IIf(primeiro, "TOTAL POR OPERAÇÃO", "&nbsp;") & _
                                    "</td></tr>"
                    End If

                    writer.Write(strHTML)

                    intLinhas += 1
                End If

                strHTML = "<tr>" & _
                            "<td colspan=""8"">" & _
                            Convert.ToInt32(drResumoNotaFiscal("Operacao")).ToString("00") & "-" & Convert.ToInt32(drResumoNotaFiscal("SubOperacao")).ToString("00") & _
                            "&nbsp;" & drResumoNotaFiscal("Descricao") & "</td>" & _
                            "<td align=""center"">" & _
                            drResumoNotaFiscal("EntradaSaida").ToString()

                If ProdutoAgrupado = False Then
                    strHTML &= "<td>" & _
                               "&nbsp;" & _
                               "</td>" & _
                               "<td align=""right"">" & _
                               Convert.ToInt32(drResumoNotaFiscal("PesoBalanca")).ToString("N0") & _
                               "</td>" & _
                               "<td align=""right"">" & _
                               Convert.ToInt32(drResumoNotaFiscal("Umidade")).ToString("N0") & _
                               "</td>" & _
                               "<td align=""right"">" & _
                               Convert.ToInt32(drResumoNotaFiscal("Impureza")).ToString("N0") & _
                               "</td>" & _
                               "<td align=""right"">" & _
                               Convert.ToInt32(drResumoNotaFiscal("Avariados")).ToString("N0") & _
                               "</td>" & _
                               "<td align=""right"">" & _
                               Convert.ToInt32(drResumoNotaFiscal("PH")).ToString("N0") & _
                               "</td>" & _
                               "<td align=""right"">" & _
                               Convert.ToInt32(drResumoNotaFiscal("GMO")).ToString("N0") & _
                               "</td>" & _
                               "<td align=""right"">" & _
                               Convert.ToInt32(drResumoNotaFiscal("PesoLiquido")).ToString("N0") & _
                               "</td>" & _
                               "<td align=""right"">" & _
                               Convert.ToInt32(drResumoNotaFiscal("PesoFiscal")).ToString("N0") & _
                               "</td>"
                Else
                    strHTML &= "<td align=""right"">" & _
                                Convert.ToInt32(drResumoNotaFiscal("PesoLiquido")).ToString("N4") & _
                                "</td>" & _
                                "<td align=""right"">" & _
                                Convert.ToInt32(drResumoNotaFiscal("PesoFiscal")).ToString("N4") & _
                                "</td>"
                End If

                strHTML &= "<td align=""right"">" &
                            "&nbsp;" &
                            "</td>" &
                            "<td align=""right"">" &
                            Convert.ToDouble(drResumoNotaFiscal("Valor")).ToString("N2") &
                            "</td>" &
                            "<td align=""right"">" &
                            Convert.ToDouble(drResumoNotaFiscal("Liquido")).ToString("N2") &
                            "</td>" &
                            "<td align=""right""><a id=""imgEncargos"" alt="""" height=""15"" width=""15"" title=""Encargos da Nota Fiscal por Operação"" style=""cursor: pointer"" onclick=""window.open('EncargosDaNota.aspx?tipo=OPNOTA&emp=" &
                            strEmpresa(0).ToString() & "&endemp=" & strEmpresa(1).ToString() &
                            "&cli=" & Pedido.CodigoCliente.ToString() &
                            "&endcli=" & Pedido.EnderecoCliente.ToString() &
                            "&ope=" & Convert.ToInt32(drResumoNotaFiscal("Operacao")).ToString("00") &
                            "&sop=" & Convert.ToInt32(drResumoNotaFiscal("SubOperacao")).ToString("00") &
                            "&pedido=" & Pedido.Codigo.ToString() & "', '', 'resizable=no, menubar=no, scrollbars=No, width=500, height=280, top=100, left=200');""><i class='fa fa-arrow-right seta'></i> </a></td></tr>"

                writer.Write(strHTML)

                intLinhas += 1

                If strDesPrd = "S" Then
                    If intLinhas = 38 Then
                        blnMudar = True
                        intLinhas = 0
                    End If
                Else
                    If intLinhas = 45 Then
                        blnMudar = True
                        intLinhas = 0
                    End If
                End If

                primeiro = False
            Next
        End If
    End Sub

#End Region

#Region "Contratos"
    ''********************************************************************************************************************
    ''********************************************************************************************************************
    ''MARCELO, COLOQUEI ESSA ROTINA NOVAMENTE PORQUE NÃO ESTAVA LISTANDO OS ITENS ESTORNO/COMPLEMENTO DO(S) CONTRATO(S)
    ''VEJA O EMAIL QUE TE MANDEI DIA 03/05, ASSIM Q AJUSTAR PODE TIRAR E VOLTAR A SUA ACIMA - FURLAN - 04/05/2010
    Private Sub ImprimirContratos(ByRef writer As HtmlTextWriter, ByVal Dados As DataSet, ByVal Pedido As [Lib].Negocio.Pedido)
        Dim dsContratos As DataSet = Dados

        If dsContratos.Tables(0).Rows.Count > 0 Then
            Dim blnMudar As Boolean = False
            Dim blnTitulo As Boolean = True
            Dim dblSomas As Double() = {0.0, 0.0, 0.0}
            Dim strHTML As String = ""

            If strDesPrd = "S" Then
                If intLinhas = 38 Then
                    blnMudar = True
                    intLinhas = 0
                    blnTitulo = False
                End If
            Else
                If intLinhas = 54 Or (intLinhas = 45 And blnTitulo) Then
                    blnMudar = True
                    intLinhas = 0
                    blnTitulo = False
                End If
            End If

            If Not blnMudar Then ImprimirCabecalhoContratos(writer)

            Dim objProduto As New [Lib].Negocio.Produto(dsContratos.Tables(0).Rows(0).Item("Produto"))
            If objProduto.Agrupar = "S" Then ProdutoAgrupado = True

            For Each drContrato As DataRow In dsContratos.Tables(0).Rows
                If blnMudar Then
                    writer.Write("</table>")
                    writer.Write("</td></tr></table>")
                    writer.Write("<br class=""quebrapagina"" />")
                    ImprimirCabecalho(writer, Pedido)
                    ImprimirCabecalhoPedido(writer, Pedido.Codigo, Pedido.CodigoOperacao, _
                                                    Pedido.CodigoSubOperacao, Pedido.SubOperacao.Descricao, _
                                                    Pedido.SubOperacao.EntradaSaida.ToString().Substring(0, 1), Pedido.CodigoSafra, Pedido.FreteCIFFOB, getTipoOperacao(Pedido), Pedido.Contrato)
                    ImprimirCabecalhoContratos(writer)
                    blnMudar = False
                End If

                strHTML = "<tr>" & _
                            "<td align=""left"">" & _
                                Convert.ToDateTime(drContrato("DataPedido")).ToString("dd/MM/yyy") & _
                            "</td>" & _
                            "<td align=""center"">" & _
                                drContrato("TipoDeLancamento").ToString() & _
                            "</td>" & _
                            "<td align=""left"">" & _
                                drContrato("NomeProduto").ToString() & _
                            "</td>" & _
                            "<td align=""right"">" & _
                                drContrato("IndiceFixado").ToString() & _
                            "</td>"

                If ProdutoAgrupado = False Then
                    strHTML &= "<td align=""right"">" & _
                                    Convert.ToInt32(drContrato("Quantidade")).ToString("N0") & _
                                "</td>"
                Else
                    strHTML &= "<td align=""right"">" & _
                                    Convert.ToDouble(drContrato("Quantidade")).ToString("N4") & _
                                "</td>"
                End If

                strHTML &= "<td align=""center"">" & _
                                drContrato("NomeMoeda").ToString() & _
                            "</td>" & _
                            "<td align=""right"">" & _
                                Convert.ToDouble(drContrato("UnitarioOficial")).ToString("N10") & _
                            "</td>" & _
                            "<td align=""right"">" & _
                                Convert.ToDouble(drContrato("UnitarioMoeda")).ToString("N10") & _
                            "</td>" & _
                            "<td align=""right"">" & _
                                Convert.ToDouble(drContrato("TotalOficial")).ToString("N2") & _
                            "</td>" & _
                            "<td align=""right"">" & _
                                Convert.ToDouble(drContrato("TotalMoeda")).ToString("N2") & _
                            "</td>" & _
                          "</tr>"

                writer.Write(strHTML)

                If drContrato("TipoDeLancamento").ToString() = "E" Then
                    dblSomas(0) -= drContrato("Quantidade")
                    dblSomas(1) -= drContrato("TotalOficial")
                    dblSomas(2) -= drContrato("TotalMoeda")
                    objQtdePedido.QuantidadeContratada -= drContrato("Quantidade")
                Else
                    dblSomas(0) += drContrato("Quantidade")
                    dblSomas(1) += drContrato("TotalOficial")
                    dblSomas(2) += drContrato("TotalMoeda")
                    objQtdePedido.QuantidadeContratada += drContrato("Quantidade")
                End If

                intLinhas += 1

                If strDesPrd = "S" Then
                    If intLinhas = 38 Then
                        blnMudar = True
                        intLinhas = 0
                        blnTitulo = False
                    End If
                Else
                    If intLinhas = 54 Or (intLinhas = 45 And blnTitulo) Then
                        blnMudar = True
                        intLinhas = 0
                        blnTitulo = False
                    End If
                End If
            Next

            ImprimirTotaisContratos(writer, dblSomas)

            writer.Write("</table>")
        End If
    End Sub

    Private Sub ImprimirCabecalhoContratos(ByRef writer As System.Web.UI.HtmlTextWriter)
        Dim strHTML As String = "<table class=""titulos"">" & _
                                "<tr>" & _
                                "<td>CONTRATOS</td>" & _
                                "</tr>" & _
                                "</table>"

        writer.Write(strHTML)

        intLinhas += 1

        strHTML = "<table class=""titulos"">" & _
                  "<tr>" & _
                  "<td class=""linhastitulobaixo"">&nbsp;<br />Data</td>" & _
                  "<td class=""linhastitulobaixo"" align=""center"">&nbsp;<br />Tipo</td>" & _
                  "<td class=""linhastitulobaixo"" align=""left"">&nbsp;<br />Produto</td>" & _
                  "<td class=""linhastitulobaixo"" align=""right"">Indice<br />Fixado</td>" & _
                  "<td class=""linhastitulobaixo"" align=""right"">&nbsp;<br />Quantidade</td>" & _
                  "<td class=""linhastitulobaixo"" align=""right"">&nbsp;<br />Moeda</td>" & _
                  "<td class=""linhastitulobaixo"" align=""right"">Unit&aacute;rio<br />R$</td>" & _
                  "<td class=""linhastitulobaixo"" align=""right"">Unit&aacute;rio<br />Outra Moeda</td>" & _
                  "<td class=""linhastitulobaixo"" align=""right"">Total<br />R$</td>" & _
                  "<td class=""linhastitulobaixo"" align=""right"">Total<br />Outra Moeda</td>" & _
                  "</tr>"

        writer.Write(strHTML)

        intLinhas += 1
    End Sub

    Private Sub ImprimirTotaisContratos(ByRef writer As System.Web.UI.HtmlTextWriter, ByVal Somas As Double())
        Dim strHTML As String = "<tr>" & _
                                "   <td colspan=""4"" class=""linhastituloacima"">T O T A I S</td>"

        If ProdutoAgrupado = False Then
            strHTML &= "   <td align=""right"" class=""linhastituloacima"">" & Convert.ToInt32(Somas(0)).ToString("N0") & "</td>"
        Else
            strHTML &= "   <td align=""right"" class=""linhastituloacima"">" & Convert.ToDouble(Somas(0)).ToString("N4") & "</td>"
        End If

        strHTML &= "   <td align=""right"" class=""linhastituloacima"">&nbsp;</td>" & _
                   "   <td align=""right"" class=""linhastituloacima"">&nbsp;</td>" & _
                   "   <td align=""right"" class=""linhastituloacima"">&nbsp;</td>" & _
                   "   <td align=""right"" class=""linhastituloacima"">" & Convert.ToDouble(Somas(1)).ToString("N2") & "</td>" & _
                   "   <td align=""right"" class=""linhastituloacima"">" & Convert.ToDouble(Somas(2)).ToString("N2") & "</td>" & _
                   "</tr>"

        writer.Write(strHTML)

        intLinhas += 1
    End Sub

#End Region

#Region "Fixações"

    Private Sub ImprimirFixacoes(ByRef writer As HtmlTextWriter, ByVal Itens As [Lib].Negocio.ListPedidoXItem, ByVal Pedido As [Lib].Negocio.Pedido)
        Dim blnMudar As Boolean = False
        Dim blnTitulo As Boolean = True
        Dim dblSomas As Double() = {0.0, 0.0, 0.0, 0.0, 0.0}
        Dim strHTML As String = ""
        Dim dblLiquido As Decimal
        Dim dblLiquidoMoeda As Decimal

        If strDesPrd = "S" Then
            If intLinhas = 38 Then
                blnMudar = True
                intLinhas = 0
                blnTitulo = False
            End If
        Else
            If intLinhas = 54 Or (intLinhas = 45 And blnTitulo) Then
                blnMudar = True
                intLinhas = 0
                blnTitulo = False
            End If
        End If

        If Not blnMudar Then ImprimirCabecalhoFixacoes(writer)

        For Each objItem As [Lib].Negocio.PedidoXItem In Itens
            For Each objFixacao As [Lib].Negocio.Fixacao In objItem.Fixacoes
                If blnMudar Then
                    writer.Write("</table>")
                    writer.Write("</td></tr></table>")
                    writer.Write("<br class=""quebrapagina"" />")
                    ImprimirCabecalho(writer, Pedido)
                    ImprimirCabecalhoPedido(writer, Pedido.Codigo, Pedido.CodigoOperacao, _
                                                    Pedido.CodigoSubOperacao, Pedido.SubOperacao.Descricao, _
                                                    Pedido.SubOperacao.EntradaSaida.ToString().Substring(0, 1), Pedido.CodigoSafra, Pedido.FreteCIFFOB, getTipoOperacao(Pedido), Pedido.Contrato)
                    ImprimirCabecalhoFixacoes(writer)
                    blnMudar = False
                End If

                dblLiquido = 0
                dblLiquidoMoeda = 0

                For Enc = 0 To objFixacao.Encargos.Count - 1
                    If objFixacao.Encargos(Enc).CodigoEncargo = "LIQUIDO" Then
                        dblLiquido += objFixacao.Encargos(Enc).ValorOficial
                        dblLiquidoMoeda += objFixacao.Encargos(Enc).ValorMoeda
                    End If
                Next

                strHTML = "<tr>" & _
                          "<td align=""left"">" & objFixacao.Movimento.ToString("dd/MM/yyy") & "</td>" & _
                          "<td align=""center"">N</td>" & _
                          "<td align=""left"">" & objFixacao.ItemPedido.CodigoProduto & "-" & objFixacao.ItemPedido.Produto.Nome & "</td>" & _
                          "<td align=""right"">" & objFixacao.Codigo.ToString() & "</td>" & _
                          "<td align=""right"">" & objFixacao.IndiceFixado.ToString() & "</td>"

                If ProdutoAgrupado = False Then
                    strHTML &= "<td align=""right"">" & objFixacao.Quantidade.ToString("N0") & "</td>"
                Else
                    strHTML &= "<td align=""right"">" & objFixacao.Quantidade.ToString("N4") & "</td>"
                End If

                strHTML &= "<td align=""center"">" & objFixacao.ItemPedido.Pedido.Moeda.Descricao & "</td>" & _
                           "<td align=""right"">" & objFixacao.UnitarioOficial.ToString("N10") & "</td>" & _
                           "<td align=""right"">" & objFixacao.UnitarioMoeda.ToString("N10") & "</td>" & _
                           "<td align=""right"">" & objFixacao.TotalOficial.ToString("N2") & "</td>" & _
                           "<td align=""right"">" & dblLiquido.ToString("N2") & "</td>" & _
                           "<td align=""right"">" & objFixacao.TotalMoeda.ToString("N2") & "</td>" & _
                           "<td align=""right"">" & dblLiquidoMoeda.ToString("N2") & "</td>" & _
                           "</tr>"

                writer.Write(strHTML)

                objQtdePedido.QuantidadeFixada += objFixacao.Quantidade
                objQtdePedido.ValorFixado += objFixacao.TotalOficial

                dblSomas(0) += objFixacao.Quantidade
                dblSomas(1) += objFixacao.TotalOficial
                dblSomas(2) += dblLiquido
                dblSomas(3) += objFixacao.TotalMoeda
                dblSomas(4) += dblLiquidoMoeda

                intLinhas += 1

                If strDesPrd = "S" Then
                    If intLinhas = 38 Then
                        blnMudar = True
                        intLinhas = 0
                        blnTitulo = False
                    End If
                Else
                    If intLinhas = 54 Or (intLinhas = 45 And blnTitulo) Then
                        blnMudar = True
                        intLinhas = 0
                        blnTitulo = False
                    End If
                End If
            Next
        Next

        ImprimirTotaisFixacoes(writer, dblSomas)

        writer.Write("</table>")
    End Sub

    Private Sub ImprimirCabecalhoFixacoes(ByRef writer As System.Web.UI.HtmlTextWriter)
        Dim strHTML As String = "<table class=""titulos"">" & _
                                "<tr>" & _
                                "<td>" & _
                                "FIXA&Ccedil;&Otilde;ES" & _
                                "</td>" & _
                                "</tr>" & _
                                "</table>"

        writer.Write(strHTML)

        intLinhas += 1

        strHTML = "<table class=""titulos"">" & _
                  "<tr>" & _
                  "<td class=""linhastitulobaixo"">" & _
                  "&nbsp;<br />Data" & _
                  "</td>" & _
                  "<td class=""linhastitulobaixo"" align=""center"">" & _
                  "&nbsp;<br />Tipo" & _
                  "</td>" & _
                  "<td class=""linhastitulobaixo"" align=""left"">" & _
                  "&nbsp;<br />Produto" & _
                  "</td>" & _
                  "<td class=""linhastitulobaixo"" align=""right"">" & _
                  "&nbsp;<br />Fixação" & _
                  "</td>" & _
                  "<td class=""linhastitulobaixo"" align=""right"">" & _
                  "Indice<br />Fixado" & _
                  "</td>" & _
                  "<td class=""linhastitulobaixo"" align=""right"">" & _
                  "&nbsp;<br />Quantidade" & _
                  "</td>" & _
                  "<td class=""linhastitulobaixo"" align=""center"">" & _
                  "&nbsp;<br />Moeda" & _
                  "</td>" & _
                  "<td class=""linhastitulobaixo"" align=""right"">" & _
                  "Unit&aacute;rio<br />R$" & _
                  "</td>" & _
                  "<td class=""linhastitulobaixo"" align=""right"">" & _
                  "Unit&aacute;rio<br />Outra Moeda" & _
                  "</td>" & _
                  "<td class=""linhastitulobaixo"" align=""right"">" & _
                  "Total<br />R$" & _
                  "</td>" & _
                  "<td class=""linhastitulobaixo"" align=""right"">" & _
                  "Líquido<br />R$" & _
                  "</td>" & _
                  "<td class=""linhastitulobaixo"" align=""right"">" & _
                  "Total<br />Outra Moeda" & _
                  "</td>" & _
                  "<td class=""linhastitulobaixo"" align=""right"">" & _
                  "Líquido<br />Outra Moeda" & _
                  "</td>" & _
                  "</tr>"

        writer.Write(strHTML)

        intLinhas += 1
    End Sub

    Private Sub ImprimirTotaisFixacoes(ByRef writer As System.Web.UI.HtmlTextWriter, ByVal Somas As Double())
        Dim strHTML As String = "<tr>" & _
                                "<td class=""linhastitulocima"" colspan=""5"">" & _
                                "T O T A I S" & _
                                "</td>" & _
                                "<td class=""linhastitulocima"" align=""right"">"

        If ProdutoAgrupado = False Then
            strHTML &= Convert.ToInt32(Somas(0)).ToString("N0")
        Else
            strHTML &= Convert.ToDouble(Somas(0)).ToString("N4")
        End If

        strHTML &= "</td>" & _
                   "<td class=""linhastitulocima"" align=""right"" colspan=""3"">" & _
                   "&nbsp;" & _
                   "</td>" & _
                   "<td class=""linhastitulocima"" align=""right"">" & _
                   Convert.ToDouble(Somas(1)).ToString("N2") & _
                   "</td>" & _
                   "<td class=""linhastitulocima"" align=""right"">" & _
                   Convert.ToDouble(Somas(2)).ToString("N2") & _
                   "</td>" & _
                   "<td class=""linhastitulocima"" align=""right"">" & _
                   Convert.ToDouble(Somas(3)).ToString("N2") & _
                   "</td>" & _
                   "<td class=""linhastitulocima"" align=""right"">" & _
                   Convert.ToDouble(Somas(4)).ToString("N2") & _
                   "</td>" & _
                   "</tr>"

        writer.Write(strHTML)

        intLinhas += 1
    End Sub

#End Region

#Region "Financeiro"

    Private Sub ImprimirFinanceiro(ByRef writer As HtmlTextWriter, ByVal Dados As DataSet, ByVal Pedido As [Lib].Negocio.Pedido)
        Dim dsFinanceiros As DataSet = Dados
        Dim dblProgramado As Double = 0, dblDescontos As Double = 0, dblJuros As Double = 0, dblLiquido As Double, dblBaixado As Double = 0
        Dim dblProgramadoMoeda As Double = 0, dblDescontosMoeda As Double = 0, dblJurosMoeda As Double = 0, dblLiquidoMoeda As Double, dblBaixadoMoeda As Double = 0
        Dim blnMudar As Boolean = False
        Dim blnTitulo As Boolean = True
        Dim dblSomas As Double() = {0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0}
        Dim strHTML As String = ""
        Dim intSinal As Integer

        If strDesPrd = "S" Then
            If intLinhas = 38 Then
                blnMudar = True
                intLinhas = 0
                blnTitulo = False
            End If
        Else
            If intLinhas = 54 Or (intLinhas = 45 And blnTitulo) Then
                blnMudar = True
                intLinhas = 0
                blnTitulo = False
            End If
        End If

        If Not blnMudar Then ImprimirCabecalhoFinanceiro(writer)

        For Each drFinanceiro As DataRow In dsFinanceiros.Tables(0).Rows
            If blnMudar Then
                writer.Write("</table>")
                writer.Write("</td></tr></table>")
                writer.Write("<br class=""quebrapagina"" />")
                ImprimirCabecalho(writer, Pedido)
                ImprimirCabecalhoPedido(writer, Pedido.Codigo, Pedido.CodigoOperacao, _
                                                Pedido.CodigoSubOperacao, Pedido.SubOperacao.Descricao, _
                                                Pedido.SubOperacao.EntradaSaida.ToString().Substring(0, 1), Pedido.CodigoSafra, Pedido.FreteCIFFOB, getTipoOperacao(Pedido), Pedido.Contrato)
                ImprimirCabecalhoFinanceiro(writer)
                blnMudar = False
            End If

            If Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada And Trim(drFinanceiro("Tipo")).ToString() = "R" Then
                intSinal = -1
            ElseIf Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Saida And Trim(drFinanceiro("Tipo")).ToString() = "P" Then
                intSinal = -1
            ElseIf Trim(drFinanceiro("Tipo")).ToString() = "RA" Then
            Else
                intSinal = 1
            End If

            dblProgramado = Convert.ToDouble(drFinanceiro("ValorDoDocumento"))
            dblDescontos = Convert.ToDouble(drFinanceiro("Descontos")) + Convert.ToDouble(drFinanceiro("Deducoes"))
            dblJuros = Convert.ToDouble(drFinanceiro("Juros")) + Convert.ToDouble(drFinanceiro("Acrescimos"))
            dblLiquido = Convert.ToDouble(drFinanceiro("ValorLiquido"))
            dblBaixado = IIf(Convert.ToDouble(drFinanceiro("Provisao")) = 1, Convert.ToDouble(drFinanceiro("ValorLiquido")), 0)

            dblProgramadoMoeda = Convert.ToDouble(drFinanceiro("MoedaValorDoDocumento"))
            dblDescontosMoeda = Convert.ToDouble(drFinanceiro("MoedaDescontos")) + Convert.ToDouble(drFinanceiro("MoedaDeducoes"))
            dblJurosMoeda = Convert.ToDouble(drFinanceiro("MoedaJuros")) + Convert.ToDouble(drFinanceiro("MoedaAcrescimos"))
            dblLiquidoMoeda = Convert.ToDouble(drFinanceiro("MoedaValorLiquido"))
            dblBaixadoMoeda = IIf(Convert.ToDouble(drFinanceiro("Provisao")) = 1, Convert.ToDouble(drFinanceiro("MoedaValorLiquido")), 0)

            strHTML = "<tr>" & _
                      "<td align=""left"">" & _
                      drFinanceiro("Tipo").ToString() & "-" & drFinanceiro("Registro").ToString() & _
                      "</td>"

            If drFinanceiro("Provisao") = 1 Then
                strHTML &= "<td align=""left"">(" & _
                      drFinanceiro("Provisao").ToString() & ")Baixa</td>"
            ElseIf drFinanceiro("Provisao") = 2 Then
                strHTML &= "<td align=""left"">(" & _
                      drFinanceiro("Provisao").ToString() & ")Previsão</td>"
            ElseIf drFinanceiro("Provisao") = 3 Then
                strHTML &= "<td align=""left"">(" & _
                      drFinanceiro("Provisao").ToString() & ")Provisão</td>"
            Else
                strHTML &= "<td align=""left"">nbsp;</td>"
            End If

            strHTML &= "<td align=""center"">" & _
                      Convert.ToDateTime(drFinanceiro("Movimento")).ToString("dd/MM/yyy") & _
                      "</td>" & _
                      "<td align=""center"">" & _
                      Convert.ToDateTime(drFinanceiro("Vencimento")).ToString("dd/MM/yyy") & _
                      "</td>" & _
                      "<td align=""center"">"
            If drFinanceiro("Provisao") = 1 OrElse drFinanceiro("Provisao") = 5 Then
                strHTML &= Convert.ToDateTime(drFinanceiro("Baixa")).ToString("dd/MM/yyy")

            End If
            strHTML &= "</td>" & _
                      "<td align=""right"">" & _
                      dblProgramado.ToString("N2") & _
                      "</td>" & _
                      "<td align=""right"">" & _
                      dblDescontos.ToString("N2") & _
                      "</td>" & _
                      "<td align=""right"">" & _
                      dblJuros.ToString("N2") & _
                      "</td>" & _
                      "<td align=""right"">" & _
                      dblLiquido.ToString("N2") & _
                      "</td>" & _
                      "<td align=""right"">" & _
                      dblBaixado.ToString("N2") & _
                      "</td>" & _
                      "<td align=""right"">" & _
                      dblProgramadoMoeda.ToString("N2") & _
                      "</td>" & _
                      "<td align=""right"">" & _
                      dblDescontosMoeda.ToString("N2") & _
                      "</td>" & _
                      "<td align=""right"">" & _
                      dblJurosMoeda.ToString("N2") & _
                      "</td>" & _
                      "<td align=""right"">" & _
                      dblLiquidoMoeda.ToString("N2") & _
                      "</td>" & _
                      "<td align=""right"">" & _
                      dblBaixadoMoeda.ToString("N2") & _
                      "</td>" & _
                      "</tr>"

            writer.Write(strHTML)

            If Not Trim(drFinanceiro("Tipo")).ToString() = "RA" Then
                If drFinanceiro("Provisao") = 1 Then objQtdePedido.ValorPago += (dblBaixado + dblDescontos - dblJuros) * intSinal

                dblSomas(0) += dblProgramado * intSinal
                dblSomas(1) += dblDescontos * intSinal
                dblSomas(2) += dblJuros * intSinal
                dblSomas(3) += dblLiquido * intSinal
                dblSomas(4) += dblBaixado * intSinal
                dblSomas(5) += dblProgramadoMoeda * intSinal
                dblSomas(6) += dblDescontosMoeda * intSinal
                dblSomas(7) += dblJurosMoeda * intSinal
                dblSomas(8) += dblLiquidoMoeda * intSinal
                dblSomas(9) += dblBaixadoMoeda * intSinal
            End If

            intLinhas += 1

            If strDesPrd = "S" Then
                If intLinhas = 38 Then
                    blnMudar = True
                    intLinhas = 0
                    blnTitulo = False
                End If
            Else
                If intLinhas = 54 Or (intLinhas = 45 And blnTitulo) Then
                    blnMudar = True
                    intLinhas = 0
                    blnTitulo = False
                End If
            End If
        Next

        ImprimirTotaisFinanceiro(writer, dblSomas)

        writer.Write("</table>")
    End Sub

    Private Sub ImprimirCabecalhoFinanceiro(ByRef writer As System.Web.UI.HtmlTextWriter)
        Dim strHTML As String = "<table class=""titulos"">" & _
                                "<tr>" & _
                                "<td>" & _
                                "FINANCEIRO - (P)Pagar - (R)Receber - (T)Troca - (RA) Recebimento de Adto." & _
                                "</td>" & _
                                "</tr>" & _
                                "</table>"

        writer.Write(strHTML)

        intLinhas += 1

        strHTML = "<table class=""titulos"">" & _
                    "<tr>" & _
                    "<td class=""linhastitulobaixo"">" & _
                    "&nbsp;<br />&nbsp;<br />Registro" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"">" & _
                    "&nbsp;<br />&nbsp;<br />Tipo" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"" align=""center"">" & _
                    "&nbsp;<br />&nbsp;<br />Movimento" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"" align=""center"">" & _
                    "&nbsp;<br />&nbsp;<br />Vencimento" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"" align=""center"">" & _
                    "&nbsp;<br />&nbsp;<br />Baixa" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"" align=""right"">" & _
                    "Valor<br />Programado<br />R$" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"" align=""right"">" & _
                    "Descontos<br />Deduções<br />R$" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"" align=""right"">" & _
                    "Juros<br />Acréscimos<br />R$" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"" align=""right"">" & _
                    "Valor<br />Líquido<br />R$" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"" align=""right"">" & _
                    "Valor<br />Baixado<br />R$" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"" align=""right"">" & _
                    "Valor<br />Programado<br />U$" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"" align=""right"">" & _
                    "Descontos<br />Deduções<br />U$" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"" align=""right"">" & _
                    "Juros<br />Acréscimos<br />U$" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"" align=""right"">" & _
                    "Valor<br />Líquido<br />U$" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"" align=""right"">" & _
                    "Valor<br />Baixado<br />U$" & _
                    "</td>" & _
                    "</tr>"

        writer.Write(strHTML)

        intLinhas += 1
    End Sub

    Private Sub ImprimirTotaisFinanceiro(ByRef writer As System.Web.UI.HtmlTextWriter, ByVal Somas As Double())
        Dim strHTML As String = "<tr>" & _
                                "<td colspan=""5"" class=""linhastitulocima"">" & _
                                "T O T A I S" & _
                                "</td>" & _
                                "<td class=""linhastitulocima"" align=""right"">" & _
                                Convert.ToDouble(Somas(0)).ToString("N2") & _
                                "</td>" & _
                                "<td class=""linhastitulocima"" align=""right"">" & _
                                Convert.ToDouble(Somas(1)).ToString("N2") & _
                                "</td>" & _
                                "<td class=""linhastitulocima"" align=""right"">" & _
                                Convert.ToDouble(Somas(2)).ToString("N2") & _
                                "</td>" & _
                                "<td class=""linhastitulocima"" align=""right"">" & _
                                Convert.ToDouble(Somas(3)).ToString("N2") & _
                                "</td>" & _
                                "<td class=""linhastitulocima"" align=""right"">" & _
                                Convert.ToDouble(Somas(4)).ToString("N2") & _
                                "</td>" & _
                                "<td class=""linhastitulocima"" align=""right"">" & _
                                Convert.ToDouble(Somas(5)).ToString("N2") & _
                                "</td>" & _
                                "<td class=""linhastitulocima"" align=""right"">" & _
                                Convert.ToDouble(Somas(6)).ToString("N2") & _
                                "</td>" & _
                                "<td class=""linhastitulocima"" align=""right"">" & _
                                Convert.ToDouble(Somas(7)).ToString("N2") & _
                                "</td>" & _
                                "<td class=""linhastitulocima"" align=""right"">" & _
                                Convert.ToDouble(Somas(8)).ToString("N2") & _
                                "</td>" & _
                                "<td class=""linhastitulocima"" align=""right"">" & _
                                Convert.ToDouble(Somas(9)).ToString("N2") & _
                                "</td>" & _
                                "</tr>"

        writer.Write(strHTML)

        intLinhas += 1
    End Sub

#End Region

#Region "Razao"
    Private Sub ImprimirRazao(ByRef writer As HtmlTextWriter, ByVal Dados As DataSet, ByVal Pedido As [Lib].Negocio.Pedido)
        Dim dsRazao As DataSet = Dados
        Dim blnMudar As Boolean = False
        Dim blnTitulo As Boolean = True
        Dim dblSomas As Double() = {0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0}
        Dim strHTML As String = ""
        'Dim intSinal As Integer
        Dim Hist1 As String = "", Hist2 As String = "", Hist3 As String = "", Hist4 As String = "", Hist5 As String = ""

        If strDesPrd = "S" Then
            If intLinhas = 38 Then
                blnMudar = True
                intLinhas = 0
                blnTitulo = False
            End If
        Else
            If intLinhas = 54 Or (intLinhas = 45 And blnTitulo) Then
                blnMudar = True
                intLinhas = 0
                blnTitulo = False
            End If
        End If

        If Not blnMudar Then ImprimirCabecalhoRazao(writer)

        For Each drRazao As DataRow In dsRazao.Tables(0).Rows
            If strDesPrd = "S" Then
                If intLinhas = 38 Then
                    blnMudar = True
                    intLinhas = 0
                    blnTitulo = False
                End If
            Else
                If intLinhas = 54 Or (intLinhas = 45 And blnTitulo) Then
                    blnMudar = True
                    intLinhas = 0
                    blnTitulo = False
                End If
            End If

            If blnMudar Then
                writer.Write("</table>")
                writer.Write("</td></tr></table>")
                writer.Write("<br class=""quebrapagina"" />")
                ImprimirCabecalho(writer, Pedido)
                ImprimirCabecalhoPedido(writer, Pedido.Codigo, Pedido.CodigoOperacao, _
                                                Pedido.CodigoSubOperacao, Pedido.SubOperacao.Descricao, _
                                                Pedido.SubOperacao.EntradaSaida.ToString().Substring(0, 1), Pedido.CodigoSafra, Pedido.FreteCIFFOB, getTipoOperacao(Pedido), Pedido.Contrato)
                ImprimirCabecalhoRazao(writer)
                blnMudar = False
            End If

            Hist1 = ""
            Hist2 = ""
            Hist3 = ""
            Hist4 = ""
            Hist5 = ""

            Hist1 = Left(drRazao("Historico"), 40)
            Hist2 = Mid(drRazao("Historico"), 41, 40)
            Hist3 = Mid(drRazao("Historico"), 81, 40)
            Hist4 = Mid(drRazao("Historico"), 121, 40)
            Hist5 = Mid(drRazao("Historico"), 161, 40)

            If Hist5.Length > 0 Then
                ImprimirInicioHistoricoRazao(writer, Hist1, Pedido, drRazao)
                ImprimirContinuacaoHistoricoRazao(writer, Hist2, Pedido, drRazao)
                ImprimirContinuacaoHistoricoRazao(writer, Hist3, Pedido, drRazao)
                ImprimirContinuacaoHistoricoRazao(writer, Hist4, Pedido, drRazao)
                ImprimirFinalHistoricoRazao(writer, Hist5, Pedido, drRazao)
            ElseIf Hist4.Length > 0 Then
                ImprimirInicioHistoricoRazao(writer, Hist1, Pedido, drRazao)
                ImprimirContinuacaoHistoricoRazao(writer, Hist2, Pedido, drRazao)
                ImprimirContinuacaoHistoricoRazao(writer, Hist3, Pedido, drRazao)
                ImprimirFinalHistoricoRazao(writer, Hist4, Pedido, drRazao)
            ElseIf Hist3.Length > 0 Then
                ImprimirInicioHistoricoRazao(writer, Hist1, Pedido, drRazao)
                ImprimirContinuacaoHistoricoRazao(writer, Hist2, Pedido, drRazao)
                ImprimirFinalHistoricoRazao(writer, Hist3, Pedido, drRazao)
            ElseIf Hist2.Length > 0 Then
                ImprimirInicioHistoricoRazao(writer, Hist1, Pedido, drRazao)
                ImprimirFinalHistoricoRazao(writer, Hist2, Pedido, drRazao)
            Else
                Dim dblDebito As Double = 0, dblCredito As Double = 0, dblDebitoMoeda As Double = 0, dblCreditoMoeda As Double = 0

                dblDebito = Convert.ToDouble(drRazao("DebitoOficial"))
                dblCredito = Convert.ToDouble(drRazao("CreditoOficial"))
                dblDebitoMoeda = Convert.ToDouble(drRazao("DebitoMoeda"))
                dblCreditoMoeda = Convert.ToDouble(drRazao("CreditoMoeda"))

                strHTML = "<tr>" &
                          "<td align=""left"">" &
                          Convert.ToDateTime(drRazao("Movimento")).ToString("dd/MM/yyy") &
                          "</td>" &
                          "<td align=""left"">" &
                          drRazao("Lote").ToString() &
                          "</td>" &
                          "<td align=""left"">" &
                          drRazao("Produto").ToString() &
                          "</td>" &
                          "<td align=""left"">" &
                          drRazao("Conta").ToString() &
                          "<a id=""imgRazao"" alt="""" height=""15"" width=""15"" title=""" & drRazao("NomeConta") & """ style=""cursor: pointer""><i class='fa fa-arrow-right seta'></i></a></td>" &
                          "<td align=""left"">" &
                          Hist1 &
                          "</td>" &
                          "<td align=""right"">" &
                          dblDebito.ToString("N2") &
                          "</td>" &
                          "<td align=""right"">" &
                          dblCredito.ToString("N2") &
                          "</td>" &
                          "<td align=""right"">" &
                          drRazao("Saldo") &
                          "</td>" &
                          "<td align=""right"">" &
                          dblDebitoMoeda.ToString("N2") &
                          "</td>" &
                          "<td align=""right"">" &
                          dblCreditoMoeda.ToString("N2") &
                          "</td>" &
                          "<td align=""right"">" &
                          drRazao("SaldoMoeda") &
                          "</td>" &
                          "</tr>"

                writer.Write(strHTML)

                intLinhas += 1

                If strDesPrd = "S" Then
                    If intLinhas = 38 Then
                        blnMudar = True
                        intLinhas = 0
                        blnTitulo = False
                    End If
                Else
                    If intLinhas = 54 Or (intLinhas = 45 And blnTitulo) Then
                        blnMudar = True
                        intLinhas = 0
                        blnTitulo = False
                    End If
                End If
            End If
        Next

        writer.Write("</table>")
    End Sub

    Private Sub ImprimirInicioHistoricoRazao(ByRef writer As System.Web.UI.HtmlTextWriter, ByVal Historico As String, ByVal Pedido As [Lib].Negocio.Pedido, ByVal drRazao As DataRow)
        Dim blnMudar As Boolean = False
        Dim blnTitulo As Boolean = True

        Dim strHTML As String = "<tr>" &
                  "<td align=""left"">" &
                  Convert.ToDateTime(drRazao("Movimento")).ToString("dd/MM/yyy") &
                  "</td>" &
                  "<td align=""left"">" &
                  drRazao("Lote").ToString() &
                  "</td>" &
                  "<td align=""left"">" &
                  drRazao("Produto").ToString() &
                  "</td>" &
                  "<td align=""left"">" &
                  drRazao("Conta").ToString() &
                  "<a id=""imgRazao"" alt="""" height=""15"" width=""15"" title=""" & drRazao("NomeConta") & """ style=""cursor: pointer""><i class='fa fa-arrow-right seta'></i></a></td>" &
                  "<td align=""left"">" &
                  Historico &
                  "</td>" &
                  "<td align=""right""></td>" &
                  "<td align=""right""></td>" &
                  "<td align=""right""></td>" &
                  "<td align=""right""></td>" &
                  "<td align=""right""></td>" &
                  "<td align=""right""></td>" &
                  "</tr>"

        writer.Write(strHTML)
        intLinhas += 1

        If strDesPrd = "S" Then
            If intLinhas = 38 Then
                blnMudar = True
                intLinhas = 0
                blnTitulo = False
            End If
        Else
            If intLinhas = 54 Or (intLinhas = 45 And blnTitulo) Then
                blnMudar = True
                intLinhas = 0
                blnTitulo = False
            End If
        End If

        If blnMudar Then
            writer.Write("</table>")
            writer.Write("</td></tr></table>")
            writer.Write("<br class=""quebrapagina"" />")
            ImprimirCabecalho(writer, Pedido)
            ImprimirCabecalhoPedido(writer, Pedido.Codigo, Pedido.CodigoOperacao, _
                                            Pedido.CodigoSubOperacao, Pedido.SubOperacao.Descricao, _
                                            Pedido.SubOperacao.EntradaSaida.ToString().Substring(0, 1), Pedido.CodigoSafra, Pedido.FreteCIFFOB, getTipoOperacao(Pedido), Pedido.Contrato)
            ImprimirCabecalhoRazao(writer)
            blnMudar = False
        End If
    End Sub

    Private Sub ImprimirContinuacaoHistoricoRazao(ByRef writer As System.Web.UI.HtmlTextWriter, ByVal Historico As String, ByVal Pedido As [Lib].Negocio.Pedido, ByVal drRazao As DataRow)
        Dim blnMudar As Boolean = False
        Dim blnTitulo As Boolean = True

        Dim strHTML As String = "<tr>" & _
                  "<td align=""left""></td>" & _
                  "<td align=""left""></td>" & _
                  "<td align=""left""></td>" & _
                  "<td align=""left""></td>" & _
                  "<td align=""left"">" & _
                  Historico & _
                  "</td>" & _
                  "<td align=""right""></td>" & _
                  "<td align=""right""></td>" & _
                  "<td align=""right""></td>" & _
                  "<td align=""right""></td>" & _
                  "<td align=""right""></td>" & _
                  "<td align=""right""></td>" & _
                  "</tr>"

        writer.Write(strHTML)
        intLinhas += 1

        If strDesPrd = "S" Then
            If intLinhas = 38 Then
                blnMudar = True
                intLinhas = 0
                blnTitulo = False
            End If
        Else
            If intLinhas = 54 Or (intLinhas = 45 And blnTitulo) Then
                blnMudar = True
                intLinhas = 0
                blnTitulo = False
            End If
        End If

        If blnMudar Then
            writer.Write("</table>")
            writer.Write("</td></tr></table>")
            writer.Write("<br class=""quebrapagina"" />")
            ImprimirCabecalho(writer, Pedido)
            ImprimirCabecalhoPedido(writer, Pedido.Codigo, Pedido.CodigoOperacao, _
                                            Pedido.CodigoSubOperacao, Pedido.SubOperacao.Descricao, _
                                            Pedido.SubOperacao.EntradaSaida.ToString().Substring(0, 1), Pedido.CodigoSafra, Pedido.FreteCIFFOB, getTipoOperacao(Pedido), Pedido.Contrato)
            ImprimirCabecalhoRazao(writer)
            blnMudar = False
        End If
    End Sub

    Private Sub ImprimirFinalHistoricoRazao(ByRef writer As System.Web.UI.HtmlTextWriter, ByVal Historico As String, ByVal Pedido As [Lib].Negocio.Pedido, ByVal drRazao As DataRow)
        Dim blnMudar As Boolean = False
        Dim blnTitulo As Boolean = True
        Dim dblDebito As Double = 0, dblCredito As Double = 0, dblDebitoMoeda As Double = 0, dblCreditoMoeda As Double = 0

        dblDebito = Convert.ToDouble(drRazao("DebitoOficial"))
        dblCredito = Convert.ToDouble(drRazao("CreditoOficial"))
        dblDebitoMoeda = Convert.ToDouble(drRazao("DebitoMoeda"))
        dblCreditoMoeda = Convert.ToDouble(drRazao("CreditoMoeda"))

        Dim strHTML As String = "<tr>" & _
                  "<td align=""left""></td>" & _
                  "<td align=""left""></td>" & _
                  "<td align=""left""></td>" & _
                  "<td align=""left""></td>" & _
                  "<td align=""left"">" & _
                  Historico & _
                  "</td>" & _
                  "<td align=""right"">" & _
                  dblDebito.ToString("N2") & _
                  "</td>" & _
                  "<td align=""right"">" & _
                  dblCredito.ToString("N2") & _
                  "</td>" & _
                  "<td align=""right"">" & _
                  drRazao("Saldo") & _
                  "</td>" & _
                  "<td align=""right"">" & _
                  dblDebitoMoeda.ToString("N2") & _
                  "</td>" & _
                  "<td align=""right"">" & _
                  dblCreditoMoeda.ToString("N2") & _
                  "</td>" & _
                  "<td align=""right"">" & _
                  drRazao("SaldoMoeda") & _
                  "</td>" & _
                  "</tr>"

        writer.Write(strHTML)
        intLinhas += 1

        If strDesPrd = "S" Then
            If intLinhas = 38 Then
                blnMudar = True
                intLinhas = 0
                blnTitulo = False
            End If
        Else
            If intLinhas = 54 Or (intLinhas = 45 And blnTitulo) Then
                blnMudar = True
                intLinhas = 0
                blnTitulo = False
            End If
        End If

        If blnMudar Then
            writer.Write("</table>")
            writer.Write("</td></tr></table>")
            writer.Write("<br class=""quebrapagina"" />")
            ImprimirCabecalho(writer, Pedido)
            ImprimirCabecalhoPedido(writer, Pedido.Codigo, Pedido.CodigoOperacao, _
                                            Pedido.CodigoSubOperacao, Pedido.SubOperacao.Descricao, _
                                            Pedido.SubOperacao.EntradaSaida.ToString().Substring(0, 1), Pedido.CodigoSafra, Pedido.FreteCIFFOB, getTipoOperacao(Pedido), Pedido.Contrato)
            ImprimirCabecalhoRazao(writer)
            blnMudar = False
        End If
    End Sub

    Private Sub ImprimirCabecalhoRazao(ByRef writer As System.Web.UI.HtmlTextWriter)
        Dim strHTML As String = "<table class=""titulos"">" & _
                                "<tr>" & _
                                "<td>" & _
                                "RAZÃO" & _
                                "</td>" & _
                                "</tr>" & _
                                "</table>"

        writer.Write(strHTML)

        intLinhas += 1

        strHTML = "<table class=""titulos"">" & _
                    "<tr>" & _
                    "<td class=""linhastitulobaixo"">" & _
                    "&nbsp;<br />Movimento" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"">" & _
                    "&nbsp;<br />Lote" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"">" & _
                    "&nbsp;<br />Produto" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"">" & _
                    "&nbsp;<br />Conta" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"">" & _
                    "&nbsp;<br />Historico" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"" align=""right"">" & _
                    "Débito<br />R$" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"" align=""right"">" & _
                    "Crédito<br />R$" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"" align=""right"">" & _
                    "Saldo<br />R$" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"" align=""right"">" & _
                    "Débito<br />U$" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"" align=""right"">" & _
                    "Crédito<br />U$" & _
                    "</td>" & _
                    "<td class=""linhastitulobaixo"" align=""right"">" & _
                    "Saldo<br />U$" & _
                    "</td>" & _
                    "</tr>"

        writer.Write(strHTML)

        intLinhas += 1
    End Sub
#End Region

#Region "Adiantamentos"

    Private Sub ImprimirAdiantamentos(ByRef writer As HtmlTextWriter, ByVal Dados As DataSet, ByVal Pedido As [Lib].Negocio.Pedido)
        Dim dsAdiantamentos As DataSet = Dados

        If dsAdiantamentos.Tables(0).Rows.Count > 0 Then
            Dim blnMudar As Boolean = False
            Dim blnTitulo As Boolean = True
            Dim blnOutraPagina As Boolean = False
            Dim dblTaxa, dblValor, dblValorMoeda, dblCorrecao, dblJuros, dblJurosMoeda, dblBaixas, dblBaixasMoeda As Double
            Dim dblSomas As Double() = {0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0}
            Dim strHTML As String = ""

            If strDesPrd = "S" Then
                If intLinhas = 38 Then
                    blnMudar = True
                    intLinhas = 0
                    blnTitulo = False
                End If
            Else
                If intLinhas = 54 Or (intLinhas = 45 And blnTitulo) Then
                    blnMudar = True
                    intLinhas = 0
                    blnTitulo = False
                End If
            End If

            If Not blnMudar Then ImprimirCabecalhoAdiantamentos(writer)

            For Each drAdiantamento As DataRow In dsAdiantamentos.Tables(0).Rows
                If blnMudar Then
                    writer.Write("</table>")
                    writer.Write("</td></tr></table>")
                    writer.Write("<br class=""quebrapagina"" />")
                    ImprimirCabecalho(writer, Pedido)
                    ImprimirCabecalhoPedido(writer, Pedido.Codigo, Pedido.CodigoOperacao, _
                                                    Pedido.CodigoSubOperacao, Pedido.SubOperacao.Descricao, _
                                                    Pedido.SubOperacao.EntradaSaida.ToString().Substring(0, 1), Pedido.CodigoSafra, Pedido.FreteCIFFOB, getTipoOperacao(Pedido), Pedido.Contrato)
                    ImprimirCabecalhoAdiantamentos(writer)
                    blnMudar = False
                End If

                dblTaxa = drAdiantamento("Taxa")
                dblValor = drAdiantamento("ValorOficial")
                dblValorMoeda = drAdiantamento("ValorMoeda")
                dblCorrecao = drAdiantamento("Correcao")
                dblBaixas = drAdiantamento("Baixas")
                dblBaixasMoeda = drAdiantamento("BaixasMoeda")

                strHTML = "<tr>" & _
                          "<td align=""left"">" & _
                          Convert.ToDateTime(drAdiantamento("Movimento")).ToString("dd/MM/yyy") & _
                          "</td>" & _
                          "<td align=""center"">" & _
                          Convert.ToDateTime(drAdiantamento("Vencimento")).ToString("dd/MM/yyy") & _
                          "</td>" & _
                          "<td align=""right"">" & _
                          drAdiantamento("Adiantamento").ToString() & _
                          "</td>" & _
                          "<td align=""right"">" & _
                          drAdiantamento("Titulo").ToString() & _
                          "</td>" & _
                          "<td align=""right"">" & _
                          drAdiantamento("Indexador").ToString() & _
                          "</td>" & _
                          "<td align=""right"">" & _
                          dblTaxa.ToString("N2") & _
                          "</td>" & _
                          "<td align=""right"">" & _
                          dblValor.ToString("N2") & _
                          "</td>" & _
                          "<td align=""right"">" & _
                          dblValorMoeda.ToString("N2") & _
                          "</td>" & _
                          "<td align=""right"">" & _
                          dblCorrecao.ToString("N2") & _
                          "</td>" & _
                          "<td align=""right"">" & _
                          dblJuros.ToString("N2") & _
                          "</td>" & _
                          "<td align=""right"">" & _
                          dblJurosMoeda.ToString("N2") & _
                          "</td>" & _
                          "</td>" & _
                          "<td align=""right"">" & _
                          dblBaixas.ToString("N2") & _
                          "</td>" & _
                          "<td align=""right"">" & _
                          dblBaixasMoeda.ToString("N2") & _
                          "</td>"

                dblSomas(0) += dblValor
                dblSomas(1) += dblValorMoeda
                dblSomas(2) += dblCorrecao
                dblSomas(3) += dblJuros
                dblSomas(4) += dblJurosMoeda
                dblSomas(5) += dblBaixas
                dblSomas(6) += dblBaixasMoeda
                dblSomas(7) += dblValor + dblJuros - dblBaixas
                dblSomas(8) += dblValorMoeda + dblJurosMoeda - dblBaixasMoeda

                strHTML &= "<td align=""right"">" & _
                          Convert.ToDouble(dblSomas(7)).ToString("N2") & _
                          "</td>" & _
                          "<td align=""right"">" & _
                          Convert.ToDouble(dblSomas(8)).ToString("N2") & _
                          "</td>" & _
                          "</tr>"

                writer.Write(strHTML)

                intLinhas += 1

                If strDesPrd = "S" Then
                    If intLinhas = 38 Then
                        blnMudar = True
                        intLinhas = 0
                        blnTitulo = False
                    End If
                Else
                    If intLinhas = 54 Or (intLinhas = 45 And blnTitulo) Then
                        blnMudar = True
                        intLinhas = 0
                        blnTitulo = False
                    End If
                End If
            Next
            ImprimirTotaisAdiantamentos(writer, dblSomas)
            writer.Write("</table>")
        End If
    End Sub

    Private Sub ImprimirCabecalhoAdiantamentos(ByRef writer As HtmlTextWriter)
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "titulos")
        writer.RenderBeginTag(HtmlTextWriterTag.Table)
        writer.RenderBeginTag(HtmlTextWriterTag.Tr)
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("ADIANTAMENTOS")
        writer.RenderEndTag()
        writer.RenderEndTag()
        writer.RenderEndTag()

        writer.AddAttribute(HtmlTextWriterAttribute.Class, "titulos")
        writer.RenderBeginTag(HtmlTextWriterTag.Table)
        writer.RenderBeginTag(HtmlTextWriterTag.Tr)
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("&nbsp;<br />Data")
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "center")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("&nbsp;<br />Vencimento")
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("N&uacute;mero<br />Adiantamento")
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("&nbsp;<br />Título")
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("&nbsp;<br />Indexador")
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("&nbsp;<br />Taxa")
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("Valor<br />Oficial")
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("Valor<br />Outra Moeda")
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("&nbsp;<br />Correção")
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("Juros<br />Oficial")
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("Juros<br />Outra Moeda")
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("Baixas<br />Oficial")
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("Baixas<br />Outra Moeda")
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("Saldo<br />Oficial")
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("Saldo<br />Outra Moeda")
        writer.RenderEndTag()
        writer.RenderEndTag()
    End Sub

    Private Sub ImprimirTotaisAdiantamentos(ByRef writer As System.Web.UI.HtmlTextWriter, ByVal Somas As Double())
        writer.RenderBeginTag(HtmlTextWriterTag.Tr)
        writer.AddAttribute(HtmlTextWriterAttribute.Colspan, "6")
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulocima")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("T O T A I S")
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulocima")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write(Convert.ToDouble(Somas(0)).ToString("N2"))
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulocima")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write(Convert.ToDouble(Somas(1)).ToString("N2"))
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulocima")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write(Convert.ToDouble(Somas(2)).ToString("N2"))
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulocima")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write(Convert.ToDouble(Somas(3)).ToString("N2"))
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulocima")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write(Convert.ToDouble(Somas(4)).ToString("N2"))
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulocima")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write(Convert.ToDouble(Somas(5)).ToString("N2"))
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulocima")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write(Convert.ToDouble(Somas(6)).ToString("N2"))
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulocima")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write(Convert.ToDouble(Somas(7)).ToString("N2"))
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulocima")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write(Convert.ToDouble(Somas(8)).ToString("N2"))
        writer.RenderEndTag()
        writer.RenderEndTag()
        'Tirei fora porque está sobrando - Furlan - 28/12/2009
        'writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulocima")
        'writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        'writer.RenderBeginTag(HtmlTextWriterTag.Td)
        'writer.Write("&nbsp;")
        'writer.RenderEndTag()
    End Sub

#End Region

#Region "Procurações"

    Private Sub ImprimirProcuracoes(ByRef writer As HtmlTextWriter, ByVal Dados As DataSet, ByVal Cedente As Boolean, ByVal Pedido As [Lib].Negocio.Pedido)
        Dim dsProcuracoes As DataSet = Dados

        If dsProcuracoes.Tables(0).Rows.Count > 0 Then
            Dim blnMudar As Boolean = False
            Dim blnTitulo As Boolean = True
            Dim blnOutraPagina As Boolean = False
            Dim intQuantidade, intFixado As Integer
            Dim dblValor, dblFixado As Double
            Dim dblSomas As Double() = {0.0, 0.0, 0.0, 0.0, 0.0, 0.0}
            Dim strHTML As String = ""

            If strDesPrd = "S" Then
                If intLinhas = 38 Then
                    blnMudar = True
                    intLinhas = 0
                    blnTitulo = False
                End If
            Else
                If intLinhas = 54 Or (intLinhas = 45 And blnTitulo) Then
                    blnMudar = True
                    intLinhas = 0
                    blnTitulo = False
                End If
            End If

            If Not blnMudar Then ImprimirCabecalhoProcuracoes(writer, Cedente)

            For Each drProcuracao As DataRow In dsProcuracoes.Tables(0).Rows
                If blnMudar Then
                    writer.Write("</table>")
                    writer.Write("</td></tr></table>")
                    writer.Write("<br class=""quebrapagina"" />")
                    ImprimirCabecalho(writer, Pedido)
                    ImprimirCabecalhoPedido(writer, Pedido.Codigo, Pedido.CodigoOperacao, _
                                                    Pedido.CodigoSubOperacao, Pedido.SubOperacao.Descricao, _
                                                    Pedido.SubOperacao.EntradaSaida.ToString().Substring(0, 1), Pedido.CodigoSafra, Pedido.FreteCIFFOB, getTipoOperacao(Pedido), Pedido.Contrato)
                    ImprimirCabecalhoProcuracoes(writer, Cedente)
                    blnMudar = False
                End If

                intQuantidade = Convert.ToInt32(drProcuracao("Quantidade"))
                If drProcuracao.Table.Columns.IndexOf("Valor") > -1 Then dblValor = Convert.ToDouble(drProcuracao("Valor"))

                intFixado = Convert.ToInt32(drProcuracao("QuantidadeFixado"))

                If drProcuracao.Table.Columns.IndexOf("ValorFixado") > -1 Then dblFixado = Convert.ToDouble(drProcuracao("ValorFixado"))

                strHTML = "<tr>" & _
                          "<td align=""left"">" & _
                          Convert.ToInt32(drProcuracao("Procuracao")).ToString("N0") & _
                          "</td>" & _
                          "<td align=""left"">" & _
                          drProcuracao("Documento").ToString() & _
                          "</td>" & _
                          "<td align=""left"">" & _
                          drProcuracao("Cliente").ToString() & _
                          "</td>"

                If Not Cedente Then
                    strHTML &= "<td align=""center"">" & _
                               "00/00/0000" & _
                               "</td>"
                End If

                strHTML &= "<td align=""right"">" & _
                           intQuantidade.ToString("N0") & _
                           "</td>"

                If Not Cedente Then
                    strHTML &= "<td align=""right"">" & _
                               dblValor.ToString("N2") & _
                               "</td>"
                End If

                strHTML &= "<td align=""right"">" & _
                            intFixado.ToString("N0") & _
                            "</td>"

                If Not Cedente Then
                    strHTML &= "<td align=""right"">" & _
                               dblFixado.ToString("N2") & _
                               "</td>"
                End If

                Dim intSaldo As Integer = intQuantidade - intFixado
                Dim dblSaldo As Double = dblValor - dblFixado

                strHTML &= "<td align=""right"">" & _
                            intSaldo.ToString("N0") & _
                            "</td>"

                If Not Cedente Then
                    strHTML &= "<td align=""right"">" & _
                               dblSaldo.ToString("N2") & _
                               "</td>"
                End If

                dblSomas(0) += intQuantidade
                dblSomas(1) += dblValor
                dblSomas(2) += intFixado
                dblSomas(3) += dblFixado
                dblSomas(4) += intSaldo
                dblSomas(5) += dblSaldo

                writer.Write(strHTML)

                intLinhas += 1

                If Cedente Then
                    objQtdePedido.QuantidadeCedente += intQuantidade
                    'objQtdeTotal.QuantidadeCedente += intQuantidade
                Else
                    objQtdePedido.QuantidadeCessionario += intQuantidade
                    'objQtdeTotal.QuantidadeCessionario += intQuantidade
                End If

                If strDesPrd = "S" Then
                    If intLinhas = 38 Then
                        blnMudar = True
                        intLinhas = 0
                        blnTitulo = False
                    End If
                Else
                    If intLinhas = 54 Or (intLinhas = 45 And blnTitulo) Then
                        blnMudar = True
                        intLinhas = 0
                        blnTitulo = False
                    End If
                End If
            Next

            If Not Cedente Then ImprimirTotaisProcuracoes(writer, dblSomas)

            writer.Write("</table>")
        End If
    End Sub

    Private Sub ImprimirCabecalhoProcuracoes(ByRef writer As HtmlTextWriter, ByVal Cedente As Boolean)
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "titulos")
        writer.RenderBeginTag(HtmlTextWriterTag.Table)
        writer.RenderBeginTag(HtmlTextWriterTag.Tr)
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("PROCURA&Ccedil;&Otilde;ES " & IIf(Cedente, "CONCEDIDAS", "OBTIDAS"))
        writer.RenderEndTag()
        writer.RenderEndTag()
        writer.RenderEndTag()

        writer.AddAttribute(HtmlTextWriterAttribute.Class, "titulos")
        writer.RenderBeginTag(HtmlTextWriterTag.Table)
        writer.RenderBeginTag(HtmlTextWriterTag.Tr)
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("&nbsp;<br />Procura&ccedil;&atilde;o")
        writer.RenderEndTag()

        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("&nbsp;<br />Documento")
        writer.RenderEndTag()

        If Cedente Then
            writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
            writer.AddAttribute(HtmlTextWriterAttribute.Align, "left")
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write("&nbsp;<br />Procurador")
            writer.RenderEndTag()
        Else
            writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
            writer.AddAttribute(HtmlTextWriterAttribute.Align, "left")
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write("&nbsp;<br />Cliente")
            writer.RenderEndTag()
            writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
            writer.AddAttribute(HtmlTextWriterAttribute.Align, "center")
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write("&nbsp;<br />Vencto.")
            writer.RenderEndTag()
        End If

        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("Peso<br />Procura&ccedil;&atilde;o")
        writer.RenderEndTag()

        If Not Cedente Then
            writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
            writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write("Valor<br />Procura&ccedil;&atilde;o")
            writer.RenderEndTag()
        End If

        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("Peso<br />Fixado")
        writer.RenderEndTag()

        If Not Cedente Then
            writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
            writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write("Valor<br />Fixado")
            writer.RenderEndTag()
        End If

        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("Saldo<br />A Fixar")
        writer.RenderEndTag()

        If Not Cedente Then
            writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulobaixo")
            writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
            writer.RenderBeginTag(HtmlTextWriterTag.Td)
            writer.Write("Valor<br />A Fixar")
            writer.RenderEndTag()
        End If

        writer.RenderEndTag()
    End Sub

    Private Sub ImprimirTotaisProcuracoes(ByRef writer As System.Web.UI.HtmlTextWriter, ByVal Somas As Double())
        writer.RenderBeginTag(HtmlTextWriterTag.Tr)
        writer.AddAttribute(HtmlTextWriterAttribute.Colspan, "4")
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulocima")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write("T O T A I S")
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulocima")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write(Convert.ToInt32(Somas(0)).ToString("N0"))
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulocima")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write(Convert.ToDouble(Somas(1)).ToString("N2"))
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulocima")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write(Convert.ToDouble(Somas(2)).ToString("N0"))
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulocima")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write(Convert.ToDouble(Somas(3)).ToString("N2"))
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulocima")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write(Convert.ToDouble(Somas(4)).ToString("N0"))
        writer.RenderEndTag()
        writer.AddAttribute(HtmlTextWriterAttribute.Class, "linhastitulocima")
        writer.AddAttribute(HtmlTextWriterAttribute.Align, "right")
        writer.RenderBeginTag(HtmlTextWriterTag.Td)
        writer.Write(Convert.ToDouble(Somas(5)).ToString("N2"))
        writer.RenderEndTag()
        writer.RenderEndTag()
    End Sub

#End Region

#Region "Totais"

    Private Sub ImprimirTotais(ByRef writer As HtmlTextWriter, ByVal Pedido As [Lib].Negocio.Pedido)
        Dim blnMudar As Boolean = False
        Dim blnTitulo As Boolean = True

        If strDesPrd = "S" Then
            If intLinhas = 38 Then
                blnMudar = True
                intLinhas = 0
                blnTitulo = False
            End If
        Else
            If intLinhas = 54 Or (intLinhas = 45 And blnTitulo) Then
                blnMudar = True
                intLinhas = 0
                blnTitulo = False
            End If
        End If

        If blnMudar Then
            writer.Write("</table>")
            writer.Write("</td></tr></table>")
            writer.Write("<br class=""quebrapagina"" />")
            ImprimirCabecalho(writer, Pedido)
            ImprimirCabecalhoPedido(writer, Pedido.Codigo, Pedido.CodigoOperacao, _
                                            Pedido.CodigoSubOperacao, Pedido.SubOperacao.Descricao, _
                                            Pedido.SubOperacao.EntradaSaida.ToString().Substring(0, 1), Pedido.CodigoSafra, Pedido.FreteCIFFOB, getTipoOperacao(Pedido), Pedido.Contrato)
            blnMudar = False
        End If

        'If intLinhas >= 0 And intLinhas < 40 Then
        '    writer.WriteBreak()
        'Else
        '    writer.Write("</td></tr></table>")
        '    writer.Write("</td></tr></table>")
        '    writer.Write("<br class=""quebrapagina"" />")
        '    ImprimirCabecalho(writer)
        '    ImprimirCabecalhoPedido(writer, Pedido.Codigo, Pedido.CodigoOperacao, _
        '                                    Pedido.CodigoSubOperacao, Pedido.SubOperacao.Descricao, _
        '                                    Pedido.SubOperacao.EntradaSaida.ToString().Substring(0, 1), Pedido.Safra)
        'End If

        Dim vlrFixado As Decimal = 0
        Dim Sql As String
        Sql = "SELECT isnull(ValorOficial,0) AS ValorOficial " & vbCrLf & _
              "  FROM VW_PedidosXItensXFixacoesXEncargos " & vbCrLf & _
              " WHERE Empresa_Id    ='" & Pedido.CodigoEmpresa & "' " & vbCrLf & _
              "   AND EndEmpresa_Id = " & Pedido.EnderecoEmpresa & vbCrLf & _
              "   AND Pedido_Id     = " & Pedido.Codigo & vbCrLf & _
              "   AND Encargo_Id    ='LIQUIDO'"

        Dim dsPedXFixXEnc As New DataSet
        dsPedXFixXEnc = Banco.ConsultaDataSet(Sql, "PedXFixXEnc")

        For Each dr As DataRow In dsPedXFixXEnc.Tables(0).Rows
            vlrFixado += dr("ValorOficial")
        Next

        'Resumo da Ficha
        'Pedido           Notas                Procuracoes        Financeiro
        'Contratado       Entregue             Cessionário        Valor Fixado
        'Fixado           Devolvido            Cedente            Valor Pago
        'Saldo à Fixar    Saldo à Entregar     Saldo              Saldo à Pagar

        Dim strHTML As String = "<table class=""titulos"">" & _
                                    "<tr>" & _
                                        "<td align=""left"" style=""font-weight: bold"">" & _
                                            "Qtde. Contratada" & _
                                        "</td>"

        If objQtdePedido.tClasse = eClassesOperacoes.AFIXAR OrElse _
           objQtdePedido.tClasse = eClassesOperacoes.DEPOSITOS Then
            If objQtdePedido.QuantidadeContratada > 0 Then
                strHTML &= "<td align=""right"" style=""font-weight: bold"">" & _
                                                    IIf(Pedido.Itens(0).Produto.Agrupar = "S", objQtdePedido.QuantidadeContratada.ToString("N4"), objQtdePedido.QuantidadeContratada.ToString("N0")) & _
                                                "</td>"
            Else
                strHTML &= "<td align=""right"" style=""font-weight: bold"">" & _
                                                    IIf(Pedido.Itens(0).Produto.Agrupar = "S", (objQtdePedido.QuantidadeEntregue - objQtdePedido.QuantidadeDevolvida).ToString("N4"), (objQtdePedido.QuantidadeEntregue - objQtdePedido.QuantidadeDevolvida).ToString("N0")) & _
                                                "</td>"
            End If
        Else
            strHTML &= "<td align=""right"" style=""font-weight: bold"">" & _
                                                IIf(Pedido.Itens(0).Produto.Agrupar = "S", objQtdePedido.QuantidadeContratada.ToString("N4"), objQtdePedido.QuantidadeContratada.ToString("N0")) & _
                                            "</td>"
        End If

        strHTML &= "<td align=""left"" style=""font-weight: bold"">" & _
                                            "&nbsp;" & _
                                        "</td>" & _
                                        "<td align=""left"" style=""font-weight: bold"">" & _
                                            "Qtde. Entregue" & _
                                        "</td>" & _
                                        "<td align=""right"" style=""font-weight: bold"">"
        If objQtdePedido.tClasse = eClassesOperacoes.MUTUO OrElse _
           objQtdePedido.tClasse = eClassesOperacoes.DEPOSITOS Then
            strHTML &= IIf(Pedido.Itens(0).Produto.Agrupar = "S", objQtdePedido.QuantidadeEntregue.ToString("N4"), objQtdePedido.QuantidadeEntregue.ToString("N0"))
        Else
            strHTML &= IIf(Pedido.Itens(0).Produto.Agrupar = "S", (objQtdePedido.QuantidadeEntregue - objQtdePedido.QuantidadeDevolvida).ToString("N4"), (objQtdePedido.QuantidadeEntregue - objQtdePedido.QuantidadeDevolvida).ToString("N0"))
        End If

        strHTML &= "</td>" & _
         "<td align=""left"" style=""font-weight: bold"">" & _
             "&nbsp;" & _
         "</td>" & _
         "<td align=""left"" style=""font-weight: bold"">" & _
             "Qtde. Cessionário" & _
         "</td>" & _
         "<td align=""right"" style=""font-weight: bold"">" & _
             IIf(Pedido.Itens(0).Produto.Agrupar = "S", objQtdePedido.QuantidadeCessionario.ToString("N4"), objQtdePedido.QuantidadeCessionario.ToString("N0")) & _
         "</td>" & _
         "<td align=""right"" style=""font-weight: bold"">" & _
             "&nbsp;" & _
         "</td>" & _
         "<td align=""left"" style=""font-weight: bold"">" & _
             "Valor Fixado" & _
         "</td>" & _
         "<td align=""right"" style=""font-weight: bold"">"

        If objQtdePedido.tClasse = eClassesOperacoes.MUTUO OrElse _
           objQtdePedido.tClasse = eClassesOperacoes.DEPOSITOS Then
            vlrFixado = 0
        End If
        strHTML &= vlrFixado.ToString("N2")

        strHTML &= "</td>" & _
        "</tr>" & _
        "<tr>" & _
        "<td align=""left"" style=""font-weight: bold"">" & _
            "Qtde. Fixada" & _
        "</td>" & _
        "<td align=""right"" style=""font-weight: bold"">"

        If objQtdePedido.tClasse = eClassesOperacoes.DEPOSITOS Then
            strHTML &= "0"
        Else
            strHTML &= IIf(Pedido.Itens(0).Produto.Agrupar = "S", objQtdePedido.QuantidadeFixada.ToString("N4"), objQtdePedido.QuantidadeFixada.ToString("N0"))
        End If

        strHTML &= "</td>" & _
        "<td align=""left"" style=""font-weight: bold"">" & _
            "&nbsp;" & _
        "</td>" & _
        "<td align=""left"" style=""font-weight: bold"">" & _
            "Qtde. Devolvida" & _
        "</td>" & _
        "<td align=""right"" style=""font-weight: bold"">" & _
            IIf(Pedido.Itens(0).Produto.Agrupar = "S", objQtdePedido.QuantidadeDevolvida.ToString("N4"), objQtdePedido.QuantidadeDevolvida.ToString("N0")) & _
        "</td>" & _
        "<td align=""left"" style=""font-weight: bold"">" & _
            "&nbsp;" & _
        "</td>" & _
        "<td align=""left"" style=""font-weight: bold"">" & _
            "Qtde. Cedente" & _
        "</td>" & _
        "<td align=""right"" style=""font-weight: bold"">" & _
            IIf(Pedido.Itens(0).Produto.Agrupar = "S", objQtdePedido.QuantidadeCedente.ToString("N4"), objQtdePedido.QuantidadeCedente.ToString("N0")) & _
        "</td>" & _
        "<td align=""left"" style=""font-weight: bold"">" & _
            "&nbsp;" & _
        "</td>" & _
        "<td align=""left"" style=""font-weight: bold"">" & _
            "Valor Pago" & _
        "</td>" & _
        "<td align=""right"" style=""font-weight: bold"">"

        If Pedido.Troca Then
            strHTML &= objQtdePedido.ValorNotaFiscal.ToString("N2")
        ElseIf Pedido.SubOperacao.Financeiro OrElse objQtdePedido.tClasse = eClassesOperacoes.AFIXAR Then
            strHTML &= objQtdePedido.ValorPago.ToString("N2")
        Else
            strHTML &= "0,00"
        End If

        strHTML &= "</td>" & _
    "</tr>" & _
    "<tr>" & _
    "<td align=""left"" style=""font-weight: bold"">" & _
        "Saldo à Fixar" & _
    "</td>"

        If objQtdePedido.tClasse = eClassesOperacoes.AFIXAR Then
            If objQtdePedido.QuantidadeContratada > 0 Then
                strHTML &= "<td align=""right"" style=""font-weight: bold"">" & _
                       IIf(Pedido.Itens(0).Produto.Agrupar = "S", (objQtdePedido.QuantidadeContratada - objQtdePedido.QuantidadeFixada).ToString("N4"), (objQtdePedido.QuantidadeContratada - objQtdePedido.QuantidadeFixada).ToString("N0")) & _
                   "</td>"
            Else
                strHTML &= "<td align=""right"" style=""font-weight: bold"">" & _
                       IIf(Pedido.Itens(0).Produto.Agrupar = "S", (objQtdePedido.QuantidadeEntregue - objQtdePedido.QuantidadeDevolvida - objQtdePedido.QuantidadeFixada).ToString("N4"), (objQtdePedido.QuantidadeEntregue - objQtdePedido.QuantidadeDevolvida - objQtdePedido.QuantidadeFixada).ToString("N0")) & _
                   "</td>"
            End If
        Else
            strHTML &= "<td align=""right"" style=""font-weight: bold"">0</td>"
        End If

        strHTML &= "<td align=""left"" style=""font-weight: bold"">" & _
                                                "&nbsp;" & _
                                            "</td>" & _
                                            "<td align=""left"" style=""font-weight: bold"">"

        If objQtdePedido.tClasse = eClassesOperacoes.MUTUO OrElse _
           objQtdePedido.tClasse = eClassesOperacoes.DEPOSITOS Then
            strHTML &= "Saldo à Devolver"
        Else
            strHTML &= "Saldo à Entregar"
        End If

        strHTML &= "</td>"

        If objQtdePedido.tClasse = eClassesOperacoes.AFIXAR Then
            If objQtdePedido.QuantidadeContratada > 0 Then
                strHTML &= "<td align=""right"" style=""font-weight: bold"">" & _
                       IIf(Pedido.Itens(0).Produto.Agrupar = "S", (objQtdePedido.QuantidadeContratada - objQtdePedido.QuantidadeEntregue + objQtdePedido.QuantidadeDevolvida).ToString("N4"), (objQtdePedido.QuantidadeContratada - objQtdePedido.QuantidadeEntregue + objQtdePedido.QuantidadeDevolvida).ToString("N0")) & _
                   "</td>"
            Else
                strHTML &= "<td align=""right"" style=""font-weight: bold"">" & _
                       IIf(Pedido.Itens(0).Produto.Agrupar = "S", (objQtdePedido.QuantidadeEntregue - objQtdePedido.QuantidadeDevolvida - objQtdePedido.QuantidadeFixada).ToString("N4"), (objQtdePedido.QuantidadeEntregue - objQtdePedido.QuantidadeDevolvida - objQtdePedido.QuantidadeFixada).ToString("N0")) & _
                   "</td>"
            End If
        ElseIf objQtdePedido.tClasse = eClassesOperacoes.MUTUO OrElse _
               objQtdePedido.tClasse = eClassesOperacoes.DEPOSITOS Then
            strHTML &= "<td align=""right"" style=""font-weight: bold"">" & _
                                                IIf(Pedido.Itens(0).Produto.Agrupar = "S", (objQtdePedido.QuantidadeEntregue - objQtdePedido.QuantidadeDevolvida).ToString("N4"), (objQtdePedido.QuantidadeEntregue - objQtdePedido.QuantidadeDevolvida).ToString("N0")) & _
                                            "</td>"
        Else
            strHTML &= "<td align=""right"" style=""font-weight: bold"">" & _
                                                IIf(Pedido.Itens(0).Produto.Agrupar = "S", (objQtdePedido.QuantidadeContratada - objQtdePedido.QuantidadeEntregue + objQtdePedido.QuantidadeDevolvida).ToString("N4"), (objQtdePedido.QuantidadeContratada - objQtdePedido.QuantidadeEntregue + objQtdePedido.QuantidadeDevolvida).ToString("N0")) & _
                                            "</td>"
        End If

        strHTML &= "<td align=""left"" style=""font-weight: bold"">" & _
                                                "&nbsp;" & _
                                            "</td>" & _
                                            "<td align=""left"" style=""font-weight: bold"">" & _
                                                "Saldo Cessão de Crédito" & _
                                            "</td>" & _
                                            "<td align=""right"" style=""font-weight: bold"">" & _
                                                IIf(Pedido.Itens(0).Produto.Agrupar = "S", (objQtdePedido.QuantidadeCedente - objQtdePedido.QuantidadeCessionario).ToString("N4"), (objQtdePedido.QuantidadeCedente - objQtdePedido.QuantidadeCessionario).ToString("N0")) & _
                                            "</td>" & _
                                            "<td align=""left"" style=""font-weight: bold"">" & _
                                                "&nbsp;" & _
                                            "</td>" & _
                                            "<td align=""left"" style=""font-weight: bold"">" & _
                                                "Saldo Financeiro" & _
                                            "</td>"
        strHTML &= "<td align=""right"" style=""font-weight: bold"">"

        If Pedido.Troca Then
            strHTML &= (vlrFixado - objQtdePedido.ValorNotaFiscal).ToString("N2")
        ElseIf Pedido.SubOperacao.Financeiro OrElse objQtdePedido.tClasse = eClassesOperacoes.AFIXAR Then
            'strHTML &= (vlrFixado - objQtdePedido.ValorPago).ToString("N2")
            strHTML &= (objQtdePedido.ValorNotaFiscal - objQtdePedido.ValorPago).ToString("N2")
        Else
            strHTML &= "0,00"
        End If

        strHTML &= "</td>" & _
                    "</tr>" & _
                    "</table>"

        writer.Write(strHTML)

        '''' VALOR LIQUIDO DA FIXAÇÃO / PESO = VLR. UNITÁRIO.... * PESO FISICO = VALOR À PAGAR

        If Pedido.SubOperacao.Financeiro = True AndAlso Pedido.Itens(0).Produto.Agrupar = "N" Then
            Dim dblLiquido As Decimal = 0
            Dim dblPeso As Decimal = 0
            For Each objItem As [Lib].Negocio.PedidoXItem In Pedido.Itens
                For Each objFixacao As [Lib].Negocio.Fixacao In objItem.Fixacoes
                    dblPeso = objFixacao.Quantidade
                    For Enc = 0 To objFixacao.Encargos.Count - 1
                        If objFixacao.Encargos(Enc).CodigoEncargo = "LIQUIDO" Then
                            dblLiquido += objFixacao.Encargos(Enc).ValorOficial
                        End If
                    Next
                Next
            Next

            Dim dblUnitario As Decimal = 0

            If dblLiquido > 0 Then dblUnitario = dblLiquido / IIf(dblPeso = 0, 1, dblPeso)

            'If objQtdePedido.QuantidadeDevolvida > 0 AndAlso objQtdePedido.QuantidadeDevolvidaFisico = 0 Then objQtdePedido.QuantidadeDevolvidaFisico = objQtdePedido.QuantidadeDevolvida
            Dim dblLiquidoNota As Decimal = 0
            If (Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada AndAlso Pedido.FreteCIFFOB = eTiposFrete.CIF) OrElse _
               (Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso Pedido.FreteCIFFOB = eTiposFrete.FOB) Then
                If objQtdePedido.QuantidadeEntregueFisico = 0 Then
                    dblLiquidoNota = Math.Round((objQtdePedido.QuantidadeEntregue - objQtdePedido.QuantidadeDevolvida) * dblUnitario, 2, MidpointRounding.AwayFromZero)
                Else
                    dblLiquidoNota = Math.Round((objQtdePedido.QuantidadeEntregueFisico - objQtdePedido.QuantidadeDevolvidaFisico) * dblUnitario, 2, MidpointRounding.AwayFromZero)
                End If
            Else
                dblLiquidoNota = Math.Round((objQtdePedido.QuantidadeEntregue - objQtdePedido.QuantidadeDevolvida) * dblUnitario, 2, MidpointRounding.AwayFromZero)
            End If

            Dim dblSaldo As Decimal = dblLiquidoNota - objQtdePedido.ValorPago

            strHTML = "<table><tr><td style=""font-weight: bold"">*** VALOR LÍQUIDO NOTA FISCAL: " & dblLiquidoNota.ToString("N2") & "&nbsp;&nbsp;&nbsp;&nbsp;</td>"
            strHTML &= "<td style=""font-weight: bold"">VALOR BAIXADO: " & objQtdePedido.ValorPago.ToString("N2") & "&nbsp;&nbsp;&nbsp;&nbsp;</td>"
            strHTML &= "<td style=""font-weight: bold"">SALDO FINANCEIRO: " & dblSaldo.ToString("N2") & "</td></tr></table>"

            writer.Write(strHTML)
        End If

        If Pedido.SubOperacao.Financeiro = True Then
            strHTML = "<table><tr><td style=""font-weight: bold"">*** ATENÇÃO, O SALDO FINANCEIRO PASSOU A SER VALOR LÍQUIDO DA NOTA FISCAL MENOS O VALOR PAGO/RECEBIDO ***</td></tr></table>"
            writer.Write(strHTML)
        End If
    End Sub

#End Region

#Region "SQL"

    Private Sub SelecionarPedidos()
        objQtdePedido.QuantidadeContratada = 0
        objQtdePedido.QuantidadeFixada = 0
        objQtdePedido.QuantidadeEntregue = 0
        objQtdePedido.QuantidadeDevolvida = 0
        objQtdePedido.QuantidadeCessionario = 0
        objQtdePedido.QuantidadeCedente = 0
        objQtdePedido.ValorFixado = 0
        objQtdePedido.ValorPago = 0

        Dim intPedido As Integer = 0
        If strPedido <> "" Then intPedido = Convert.ToInt32(strPedido)

        Dim strEmpresa As String() = strEmpresaQry.Replace(";", "-").Split("-")
        Dim intEndEmpresa As Integer = 0
        If strEmpresa.Length > 1 Then intEndEmpresa = Convert.ToInt32(strEmpresa(1))

        Dim strCliente As String() = strClienteQry.Replace(";", "-").Split("-")
        Dim intEndCliente As Integer = 0
        If strCliente.Length > 1 Then intEndCliente = Convert.ToInt32(strCliente(1))

        Dim objEntradaSaida As eEntradaSaidaNenhum = eEntradaSaidaNenhum.Nenhum

        'If blnEntrada And Not blnSaida Then
        '    objEntradaSaida = eEntradaSaidaNenhum.Entrada
        'ElseIf blnSaida And Not blnEntrada Then
        '    objEntradaSaida = eEntradaSaidaNenhum.Saida
        'End If

        objPedidos.Selecionar(CodigoPedido:=intPedido, Empresa:=strEmpresa(0), EnderecoEmpresa:=intEndEmpresa, _
                              Cliente:=strCliente(0), EnderecoCliente:=intEndCliente, Situacao:=eSituacao.Normal, _
                              EntradaSaida:=objEntradaSaida, Agrupar:=False, DataPedidoInicial:=Convert.ToDateTime(strDataFim), _
                              OperadorDataPedido:=[Lib].Negocio.eOperadoresSelecao.MenorOuIgualQue, CodigoProduto:=strProduto, Safra:=strSafra, OrdenarPor:="DataPedido")
    End Sub

    Private Function SelecionarRazao(ByVal Empresa As String, ByVal Endereco As Integer, ByVal Pedido As Integer) As DataSet
        Dim dsRazao As DataSet = Extrato.getDataSetRazao(Empresa, Endereco, Pedido, "", 0, SituacaoPedido)

        dsRazao.Tables(0).Columns.Add("Saldo", Type.GetType("System.String"))
        dsRazao.Tables(0).Columns.Add("SaldoMoeda", Type.GetType("System.String"))

        Dim Saldo As Decimal = 0
        Dim SaldoMoeda As Decimal = 0
        For Each dr As DataRow In dsRazao.Tables(0).Rows
            Saldo += dr("DebitoOficial") - dr("CreditoOficial")
            SaldoMoeda += dr("DebitoMoeda") - dr("CreditoMoeda")

            If Saldo > 0 Then
                dr("Saldo") = Saldo.ToString("N2") & "-D"
            End If
            If SaldoMoeda > 0 Then
                dr("SaldoMoeda") = SaldoMoeda.ToString("N2") & "-D"
            End If

            If Saldo < 0 Then
                dr("Saldo") = (Saldo * -1).ToString("N2") & "-C"
            End If
            If SaldoMoeda < 0 Then
                dr("SaldoMoeda") = (SaldoMoeda * -1).ToString("N2") & "-C"
            End If

            If Saldo = 0 Then
                dr("Saldo") = Saldo.ToString("N2") & "DC"
            End If
            If SaldoMoeda = 0 Then
                dr("SaldoMoeda") = SaldoMoeda.ToString("N2") & "DC"
            End If
        Next

        Return dsRazao
    End Function
#End Region

#End Region

End Class