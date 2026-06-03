Imports System.Xml
Imports System.Data
Imports System.Web
Imports System.Security.Cryptography.X509Certificates
Imports Microsoft.VisualBasic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class Pamcard

    Private Sub SetValues(ByVal key As String, ByVal value As Object, ByVal wsfields As WSPamcard.FieldTO(), ByVal Pos As Integer)
        Dim wsfield As New WSPamcard.FieldTO()
        wsfield.key = key
        wsfield.value = value
        wsfields.SetValue(wsfield, Pos)
        wsfield = Nothing
        'Pos = Pos + 1
    End Sub

    Private Function GetValue(ByVal key As String, ByVal wsfields As WSPamcard.FieldTO()) As Object
        For Each objField As WSPamcard.FieldTO In wsfields
            If objField.key = key Then Return objField.value
        Next

        Return Nothing
    End Function

    Private Sub SetDataset(ByVal Codigo As String, ByVal Valor As String, ByVal ds As DataSet)
        Dim drRow As DataRow = ds.Tables(0).NewRow()
        drRow("Codigo") = Codigo
        drRow("Valor") = Valor
        ds.Tables(0).Rows.Add(drRow)
    End Sub

    Public Function InclusaoContratoDeFrete(ByVal Empresa As String, ByVal Contrato As Negocio.NotaFiscal, ByVal objNota As Negocio.NotaFiscal, ByVal ProprietarioDoCartao As Cliente) As DataSet
        Dim ds As New DataSet
        Dim dsRet As New DataSet

        Try
            Dim dt As New DataTable("Pamcard")
            dt.Columns.Add("Codigo", Type.GetType("System.String"))
            dt.Columns.Add("Valor", Type.GetType("System.String"))
            ds.Tables.Add(dt)

            SetDataset("viagem.contratante.documento.numero", Empresa, ds)
            SetDataset("viagem.id.cliente", Contrato.Codigo, ds)
            SetDataset("viagem.contrato.numero", Contrato.Codigo, ds)
            SetDataset("viagem.favorecido.qtde", 1, ds)

            If Contrato.CartaoNovo Then
                SetDataset("viagem.favorecido1.tipo", IIf(Contrato.CodigoTransportador.Length = 11, 2, 1), ds)
                SetDataset("viagem.favorecido1.documento.qtde", IIf(Contrato.CodigoTransportador.Length = 11, 3, 2), ds)
                SetDataset("viagem.favorecido1.documento1.tipo", 2, ds)
                SetDataset("viagem.favorecido1.documento1.numero", ProprietarioDoCartao.Codigo, ds)
                SetDataset("viagem.favorecido1.documento2.tipo", 3, ds)
                SetDataset("viagem.favorecido1.documento2.numero", ProprietarioDoCartao.RG, ds)
                SetDataset("viagem.favorecido1.documento2.uf", ProprietarioDoCartao.CodigoEstado, ds)
                If Contrato.CodigoTransportador.Length = 11 Then
                    SetDataset("viagem.favorecido1.documento3.tipo", 5, ds)
                    SetDataset("viagem.favorecido1.documento3.numero", ProprietarioDoCartao.RNTRCTransportador, ds)
                End If

                SetDataset("viagem.favorecido1.nome", ProprietarioDoCartao.Nome, ds)
                SetDataset("viagem.favorecido1.data.nascimento", String.Format("{0:dd/MM/yyyy}", ProprietarioDoCartao.NascimentoConstituicao), ds)
                SetDataset("viagem.favorecido1.endereco.logradouro", ProprietarioDoCartao.Endereco, ds)
                SetDataset("viagem.favorecido1.endereco.numero", IIf(ProprietarioDoCartao.Numero.ToString.Length = 0, 0, ProprietarioDoCartao.Numero), ds)
                SetDataset("viagem.favorecido1.endereco.bairro", ProprietarioDoCartao.Bairro, ds)
                SetDataset("viagem.favorecido1.endereco.cidade.ibge", ProprietarioDoCartao.Municipio.EstadoIbge & ProprietarioDoCartao.CodigoMunicipio.ToString("00000"), ds)
                SetDataset("viagem.favorecido1.endereco.cep", Funcoes.OnlyNumbers(ProprietarioDoCartao.CEP), ds)
                SetDataset("viagem.favorecido1.telefone.ddd", String.Format("0{0}", ProprietarioDoCartao.Telefone.Substring(0, 2)), ds)
                SetDataset("viagem.favorecido1.telefone.numero", Funcoes.OnlyNumbers(ProprietarioDoCartao.Telefone).Substring(2, 8), ds)

                'Meio de Pagamento - 1 Cartão  2 - Conta Depósito
                SetDataset("viagem.favorecido1.meio.pagamento", 1, ds)

                'Número do Cartão
                SetDataset("viagem.favorecido1.cartao.numero", Contrato.CartaoPgtoFrete, ds)

                If Contrato.CodigoTransportador.Length > 11 Then
                    SetDataset("viagem.favorecido1.empresa.nome", Contrato.Transportador.Nome, ds)
                    SetDataset("viagem.favorecido1.empresa.cnpj", Contrato.Transportador.Codigo, ds)
                    SetDataset("viagem.favorecido1.empresa.rntrc", Contrato.Transportador.RNTRCTransportador, ds)
                End If
            Else
                SetDataset("viagem.favorecido1.tipo", IIf(Contrato.CodigoTransportador.Length = 11, 2, 1), ds)
                SetDataset("viagem.favorecido1.documento.qtde", IIf(Contrato.CodigoTransportador.Length = 11, 3, 2), ds)
                If Contrato.CodigoTransportador.Length = 11 Then
                    SetDataset("viagem.favorecido1.documento1.tipo", 2, ds)
                Else
                    SetDataset("viagem.favorecido1.documento1.tipo", 1, ds)
                End If
                SetDataset("viagem.favorecido1.documento1.numero", Contrato.CodigoTransportador, ds)
                If Contrato.CodigoTransportador.Length = 11 Then
                    SetDataset("viagem.favorecido1.documento2.tipo", 5, ds)
                Else
                    SetDataset("viagem.favorecido1.documento2.tipo", 6, ds)
                End If
                SetDataset("viagem.favorecido1.documento2.numero", Contrato.Transportador.RNTRCTransportador, ds)
                If Contrato.CodigoTransportador.Length = 11 Then
                    SetDataset("viagem.favorecido1.documento3.tipo", 3, ds)
                    SetDataset("viagem.favorecido1.documento3.numero", Contrato.Transportador.RG, ds)
                End If

                'Meio de Pagamento - 1 Cartão  2 - Conta Depósito
                SetDataset("viagem.favorecido1.meio.pagamento", 1, ds)

                'Número do Cartão
                SetDataset("viagem.favorecido1.cartao.numero", Contrato.CartaoPgtoFrete, ds)
            End If

            'Veículos
            Dim j As Integer = 1
            If Contrato.PlacaDetalhes.Placa02.Length > 0 Then j += 1
            If Contrato.PlacaDetalhes.Placa03.Length > 0 Then j += 1
            If Contrato.PlacaDetalhes.Placa04.Length > 0 Then j += 1

            SetDataset("viagem.veiculo.qtde", j, ds)
            SetDataset("viagem.veiculo1.placa", Contrato.PlacaDetalhes.Placa01.Replace("-", ""), ds)
            SetDataset("viagem.veiculo1.rntrc", Contrato.PlacaDetalhes.RNTRCPlaca01, ds)

            If Contrato.PlacaDetalhes.Placa02.Length > 0 Then
                SetDataset("viagem.veiculo2.placa", Contrato.PlacaDetalhes.Placa02.Replace("-", ""), ds)
                SetDataset("viagem.veiculo2.rntrc", Contrato.PlacaDetalhes.RNTRCPlaca02, ds)
            End If
            If Contrato.PlacaDetalhes.Placa03.Length > 0 Then
                SetDataset("viagem.veiculo3.placa", Contrato.PlacaDetalhes.Placa03.Replace("-", ""), ds)
                SetDataset("viagem.veiculo3.rntrc", Contrato.PlacaDetalhes.RNTRCPlaca03, ds)
            End If
            If Contrato.PlacaDetalhes.Placa04.Length > 0 Then
                SetDataset("viagem.veiculo4.placa", Contrato.PlacaDetalhes.Placa04.Replace("-", ""), ds)
                SetDataset("viagem.veiculo4.rntrc", Contrato.PlacaDetalhes.RNTRCPlaca04, ds)
            End If

            SetDataset("viagem.veiculo.categoria", Contrato.PlacaDetalhes.TipoDeVeiculoDetalhes.CodigoPamcard, ds)

            'Início e Término da Viagem
            SetDataset("viagem.data.partida", Contrato.Movimento.ToString("dd/MM/yyyy"), ds)
            SetDataset("viagem.data.termino", Contrato.DataTermino.ToString("dd/MM/yyyy"), ds)

            'Origem / Destino
            SetDataset("viagem.origem.cidade.ibge", Contrato.Deposito.Municipio.EstadoIbge & Contrato.Deposito.CodigoMunicipio.ToString("00000"), ds)
            SetDataset("viagem.destino.cidade.ibge", Contrato.Destino.Municipio.EstadoIbge & Contrato.Destino.CodigoMunicipio.ToString("00000"), ds)

            Dim vlrPedagio As Decimal = 0
            Dim vlrAdto As Decimal = 0
            Dim vlrLiquido As Decimal = 0
            Dim vlrFrete As Decimal = 0
            Dim vlrSeguro As Decimal = 0
            Dim vlrCadastro As Decimal = 0

            For Each enc As NotaFiscalXItemXEncargo In Contrato.Itens(0).Encargos
                If enc.Codigo = "TARIFA PEDAGIO" Then
                    vlrPedagio = enc.Valor
                ElseIf enc.Codigo = "ADTODEFRETE" OrElse enc.Codigo = "ADIANTAMENTO" Then
                    vlrAdto += enc.Valor
                ElseIf enc.Codigo = "TARIFA SEGURO" Then
                    vlrSeguro = enc.Valor
                ElseIf enc.Codigo = "TAXA CADASTRO" Then
                    vlrCadastro = enc.Valor
                ElseIf enc.Codigo = "PRODUTO" Then
                    vlrFrete = enc.Valor
                ElseIf enc.Codigo = "LIQUIDO" Then
                    vlrLiquido = enc.Valor
                End If
            Next

            'Pedágio - Tipo 5 PAMCARD
            If vlrPedagio > 0 Then
                SetDataset("viagem.pedagio.solucao.id", 5, ds)
                SetDataset("viagem.pedagio.valor", vlrPedagio.ToString("N2").Replace(".", "").Replace(",", "."), ds)
                SetDataset("viagem.pedagio.status.id", 2, ds)
            End If

            'Natureza Carga
            SetDataset("viagem.carga.natureza", objNota.Itens(0).Produto.DetalhesGenero.CodigoPamcard, ds)
            SetDataset("viagem.carga.peso", Contrato.PesoLiquido.ToString("N2").Replace(".", "").Replace(",", "."), ds)

            'Dados do Conhecimento
            SetDataset("viagem.documento.qtde", 1, ds)
            SetDataset("viagem.documento1.tipo", 5, ds)
            SetDataset("viagem.documento1.numero", Contrato.Codigo, ds)
            SetDataset("viagem.documento1.quantidade", Contrato.Itens(0).QuantidadeFiscal.ToString("N2").Replace(".", "").Replace(",", "."), ds)

            SetDataset("viagem.documento1.pessoafiscal.qtde", 2, ds)

            SetDataset("viagem.documento1.pessoafiscal1.tipo", 1, ds)
            SetDataset("viagem.documento1.pessoafiscal1.documento.tipo", 1, ds)
            SetDataset("viagem.documento1.pessoafiscal1.documento.numero", objNota.CodigoEmpresa, ds)
            SetDataset("viagem.documento1.pessoafiscal1.nome", objNota.Empresa.Nome, ds)
            SetDataset("viagem.documento1.pessoafiscal1.endereco.logradouro", Funcoes.EliminarCaracteresEspeciais(objNota.Empresa.Endereco), ds)
            SetDataset("viagem.documento1.pessoafiscal1.endereco.numero", objNota.Empresa.Numero, ds)
            SetDataset("viagem.documento1.pessoafiscal1.endereco.bairro", objNota.Empresa.Bairro, ds)
            SetDataset("viagem.documento1.pessoafiscal1.endereco.cidade.ibge", objNota.Empresa.Municipio.EstadoIbge & objNota.Empresa.CodigoMunicipio.ToString("00000"), ds)
            SetDataset("viagem.documento1.pessoafiscal1.endereco.cep", Funcoes.OnlyNumbers(objNota.Empresa.CEP), ds)

            SetDataset("viagem.documento1.pessoafiscal2.tipo", 2, ds)
            If objNota.CodigoCliente.Length = 11 Then
                SetDataset("viagem.documento1.pessoafiscal2.documento.tipo", 2, ds)
            Else
                SetDataset("viagem.documento1.pessoafiscal2.documento.tipo", 1, ds)
            End If
            SetDataset("viagem.documento1.pessoafiscal2.documento.numero", objNota.CodigoCliente, ds)
            SetDataset("viagem.documento1.pessoafiscal2.nome", objNota.Cliente.Nome, ds)
            SetDataset("viagem.documento1.pessoafiscal2.endereco.logradouro", Funcoes.EliminarCaracteresEspeciais(objNota.Cliente.Endereco), ds)
            SetDataset("viagem.documento1.pessoafiscal2.endereco.numero", objNota.Cliente.Numero, ds)
            SetDataset("viagem.documento1.pessoafiscal2.endereco.bairro", objNota.Cliente.Bairro, ds)
            SetDataset("viagem.documento1.pessoafiscal2.endereco.cidade.ibge", objNota.Cliente.Municipio.EstadoIbge & objNota.Cliente.CodigoMunicipio.ToString("00000"), ds)
            SetDataset("viagem.documento1.pessoafiscal2.endereco.cep", Funcoes.OnlyNumbers(objNota.Cliente.CEP), ds)

            'Dados Pagamento/Financeiro
            If vlrAdto > 0 Then
                SetDataset("viagem.parcela.qtde", 2, ds)
                'Saldo
                SetDataset("viagem.parcela1.efetivacao.tipo", 2, ds) ' Tipo 2 - Automático
                SetDataset("viagem.parcela1.valor", CDec(vlrLiquido - vlrAdto).ToString("N2").Replace(".", "").Replace(",", "."), ds)
                SetDataset("viagem.parcela1.subtipo", 3, ds) ' Tipo 3 - Saldo Final
                SetDataset("viagem.parcela1.status.id", 1, ds) ' Tipo 1 - Pendente
                SetDataset("viagem.parcela1.data", Contrato.DataTermino.ToString("dd/MM/yyyy"), ds)
                SetDataset("viagem.parcela1.favorecido.tipo.id", 1, ds)
                SetDataset("viagem.parcela1.numero.cliente", 1, ds)
                'Adiantamento
                SetDataset("viagem.parcela2.efetivacao.tipo", 2, ds) ' Tipo 2 - Automático
                SetDataset("viagem.parcela2.valor", vlrAdto.ToString("N2").Replace(".", "").Replace(",", "."), ds)
                SetDataset("viagem.parcela2.subtipo", 1, ds) ' Tipo 1 - Adiantamento
                SetDataset("viagem.parcela2.status.id", 2, ds) ' Tipo 2 - Liberado
                SetDataset("viagem.parcela2.data", Contrato.Movimento.ToString("dd/MM/yyyy"), ds)
                SetDataset("viagem.parcela2.favorecido.tipo.id", 1, ds)
                SetDataset("viagem.parcela2.numero.cliente", 2, ds)
            Else
                SetDataset("viagem.parcela.qtde", 1, ds)

                SetDataset("viagem.parcela1.efetivacao.tipo", 2, ds) ' Tipo 2 - Automático
                SetDataset("viagem.parcela1.valor", CDec(Contrato.TotalNota - vlrAdto).ToString("N2").Replace(".", "").Replace(",", "."), ds)
                SetDataset("viagem.parcela1.subtipo", 3, ds) ' Tipo 3 - Saldo Final
                SetDataset("viagem.parcela1.status.id", 1, ds) ' Tipo 1 - Pendente
                SetDataset("viagem.parcela1.data", Contrato.DataTermino.ToString("dd/MM/yyyy"), ds)
                SetDataset("viagem.parcela1.favorecido.tipo.id", 1, ds)
                SetDataset("viagem.parcela1.numero.cliente", 1, ds)
            End If

            SetDataset("viagem.frete.valor.bruto", vlrFrete.ToString("N2").Replace(".", "").Replace(",", "."), ds)

            j = 0
            If vlrSeguro > 0 Then j += 1
            If vlrCadastro > 0 Then j += 1
            SetDataset("viagem.frete.item.qtde", j, ds)

            j = 1

            If vlrSeguro > 0 Then
                SetDataset("viagem.frete.item" & j & ".tipo", 82, ds)
                SetDataset("viagem.frete.item" & j & ".valor", vlrSeguro.ToString("N2").Replace(".", "").Replace(",", "."), ds)
                j += 1
            End If

            If vlrCadastro > 0 Then
                SetDataset("viagem.frete.item" & j & ".tipo", 82, ds)
                SetDataset("viagem.frete.item" & j & ".valor", vlrCadastro.ToString("N2").Replace(".", "").Replace(",", "."), ds)
            End If

            Dim strXML As String

            strXML = "<insertContrato>"
            For Each dr As DataRow In ds.Tables(0).Rows
                strXML &= "<" & dr("Codigo") & ">" & dr("Valor") & "</" & dr("Codigo") & ">"
            Next
            strXML &= "</insertContrato>"

            Dim strNomeArquivo As String = HttpContext.Current.Server.MapPath(String.Format("Pamcard/Envio/insertContrato_{0}_{1}.xml", DateTime.Now().ToString("yyyy-MM-dd"), Contrato.Codigo))
            Dim xmlArquivo As New XmlDocument()
            xmlArquivo.LoadXml(strXML)
            xmlArquivo.Save(strNomeArquivo)
            xmlArquivo = Nothing

            Dim wsFields(ds.Tables(0).Rows.Count - 1) As WSPamcard.FieldTO
            Dim i As Integer
            For i = 0 To ds.Tables(0).Rows.Count - 1
                Dim teste As String = i
                SetValues(ds.Tables(0).Rows(i).Item("Codigo"), ds.Tables(0).Rows(i).Item("Valor"), wsFields, i)
            Next

            Dim wsRequest As New WSPamcard.RequestTO()
            wsRequest.certificate = X509Certificate.CreateFromCertFile(HttpContext.Current.Server.MapPath("Pamcard/pamcard_alvorada_04854422000185.crt")).GetRawCertData()
            wsRequest.context = "InsertFreightContract"
            wsRequest.fields = wsFields

            Dim wsPamCard As New WSPamcard.WSPamcardService()
            Dim wsResponse As New WSPamcard.ResponseTO()
            wsResponse = wsPamCard.execute(wsRequest)

            Dim wsFieldsResponse() As WSPamcard.FieldTO = wsResponse.fields

            Dim intPedagio As Integer = GetValue("viagem.pedagio.qtde", wsFieldsResponse)
            Dim intPonto As Integer = GetValue("viagem.ponto.qtde", wsFieldsResponse)

            strXML = "<insertContrato>"
            strXML &= String.Format("<erro>{0}</erro><errodescricao>{1}</errodescricao><qtdePedagio>{2}</qtdePedagio>", _
                                    GetValue("mensagem.codigo", wsFieldsResponse), _
                                    GetValue("mensagem.descricao", wsFieldsResponse), _
                                    intPedagio)

            For intCartao As Integer = 1 To intPedagio
                strXML &= "<praca>"
                strXML &= String.Format("<pSeq>{0}</pSeq><pNome>{1}</pNome><pValor>{2}</pValor>", _
                                        GetValue("viagem.pedagio.praca" & intCartao & ".seq", wsFieldsResponse), _
                                        GetValue("viagem.pedagio.praca" & intCartao & ".nome", wsFieldsResponse), _
                                        GetValue("viagem.pedagio.praca" & intCartao & ".valor", wsFieldsResponse))
                strXML &= "</praca>"
            Next

            strXML &= String.Format("<pedagioKm>{0}</pedagioKm><rotaNome>{1}</rotaNome><origemCidade>{2}</origemCidade><destinoCidade>{3}</destinoCidade>", _
                                    GetValue("viagem.pedagio.km", wsFieldsResponse), _
                                    GetValue("viagem.rota.nome", wsFieldsResponse), _
                                    GetValue("viagem.origem.cidade.ibge", wsFieldsResponse), _
                                    GetValue("viagem.destino.cidade.ibge", wsFieldsResponse))

            strXML &= "<pontoQdte>" & intPonto & "</pontoQdte>"

            For intConta As Integer = 1 To intPonto
                strXML &= "<ponto>"
                strXML &= String.Format("<pCidade>{0}</pCidade>", _
                                        GetValue("viagem.ponto" & intConta & ".cidade.ibge", wsFieldsResponse))
                strXML &= "</ponto>"
            Next

            strXML &= String.Format("<anttNumero>{0}</anttNumero><anttProtocolo>{1}</anttProtocolo><viagemId>{2}</viagemId>", _
                                    GetValue("viagem.antt.ciot.numero", wsFieldsResponse), _
                                    GetValue("viagem.antt.protocolo", wsFieldsResponse), _
                                    GetValue("viagem.id", wsFieldsResponse))

            strXML &= "</insertContrato>"

            'strXML = "<InsertFreightContract>"
            'strXML &= String.Format("<erro>{0}</erro><errodescricao>{1}</errodescricao><qtdePedagio>{2}</qtdePedagio>", _
            '                        "0", _
            '                        "TUDO CERTO", _
            '                        0)
            'strXML &= String.Format("<anttNumero>{0}</anttNumero><anttProtocolo>{1}</anttProtocolo><viagemId>{2}</viagemId>", _
            '                        "252627282930", _
            '                        "1234", _
            '                        "1065859")

            'strXML &= "</InsertFreightContract>"

            strNomeArquivo = HttpContext.Current.Server.MapPath(String.Format("Pamcard/Retorno/respInsertContrato_{0}_{1}.xml", DateTime.Now().ToString("yyyy-MM-dd"), Contrato.Codigo))

            xmlArquivo = New XmlDocument()
            xmlArquivo.LoadXml(strXML)
            xmlArquivo.Save(strNomeArquivo)
            xmlArquivo = Nothing

            dsRet.ReadXml(strNomeArquivo)

            Return dsRet
        Catch ex As Exception
            HttpContext.Current.Session("ssMessage") = Funcoes.EliminarCaracteresEspeciais(ex.Message.ToString)
            dsRet = Nothing
            Return dsRet
        End Try
    End Function

    Public Function AtualizarValoresSaldoContratoDeFrete(ByVal Empresa As String, ByVal Contrato As NotaFiscal, ByVal DataDaBaixa As String, ByVal Desconto As String, ByVal Deducoes As String, ByVal Juros As String, ByVal Acrescimo As String, ByVal ValorLiquido As String) As DataSet
        Dim ds As New DataSet
        Dim dsRet As New DataSet

        Try
            Dim tbP As New DataTable("Pamcard")
            tbP.Columns.Add("Codigo", Type.GetType("System.String"))
            tbP.Columns.Add("Valor", Type.GetType("System.String"))
            ds.Tables.Add(tbP)

            SetDataset("viagem.contratante.documento.numero", Empresa, ds)
            SetDataset("viagem.id", Contrato.idPamcard, ds)
            SetDataset("viagem.antt.ciot.numero", Contrato.ProtocoloANTT, ds)
            SetDataset("viagem.parcela.qtde", 1, ds)

            SetDataset("viagem.parcela1.efetivacao.tipo", 2, ds) ' Tipo 2 - Automático
            SetDataset("viagem.parcela1.valor", ValorLiquido.Replace(".", "").Replace(",", "."), ds)
            SetDataset("viagem.parcela1.subtipo", 3, ds) ' Tipo 3 - Saldo Final
            SetDataset("viagem.parcela1.status.id", 2, ds) ' Tipo 2 - Liberado
            SetDataset("viagem.parcela1.data", DataDaBaixa, ds)
            SetDataset("viagem.parcela1.favorecido.tipo.id", 1, ds)
            SetDataset("viagem.parcela1.numero.cliente", 1, ds)

            SetDataset("viagem.frete.valor.bruto", Contrato.TotalProduto.ToString("N2").Replace(".", "").Replace(",", "."), ds)

            Dim i As Integer = 0
            Dim j As Integer = 1

            If Desconto > 0 Then i += 1
            If Deducoes > 0 Then i += 1
            If Juros > 0 Then i += 1
            If Acrescimo > 0 Then i += 1

            If i > 0 Then
                SetDataset("viagem.frete.item.qtde", i, ds)

                If Desconto > 0 Then
                    SetDataset("viagem.frete.item" & j & ".tipo", 161, ds)
                    SetDataset("viagem.frete.item" & j & ".valor", Desconto.Replace(".", "").Replace(",", "."), ds)
                    j += 1
                End If

                If Deducoes > 0 Then
                    SetDataset("viagem.frete.item" & j & ".tipo", 204, ds)
                    SetDataset("viagem.frete.item" & j & ".valor", Deducoes.Replace(".", "").Replace(",", "."), ds)
                    j += 1
                End If

                If Juros > 0 Then
                    SetDataset("viagem.frete.item" & j & ".tipo", 5, ds)
                    SetDataset("viagem.frete.item" & j & ".valor", Juros.Replace(".", "").Replace(",", "."), ds)
                    j += 1
                End If

                If Acrescimo > 0 Then
                    SetDataset("viagem.frete.item" & j & ".tipo", 6, ds)
                    SetDataset("viagem.frete.item" & j & ".valor", Acrescimo.Replace(".", "").Replace(",", "."), ds)
                End If
            End If

            Dim strXML As String

            strXML = "<updateValuesFreightContract>"
            For Each dr As DataRow In ds.Tables(0).Rows
                strXML &= "<" & dr("Codigo") & ">" & dr("Valor") & "</" & dr("Codigo") & ">"
            Next
            strXML &= "</updateValuesFreightContract>"

            Dim strNomeArquivo As String = HttpContext.Current.Server.MapPath(String.Format("Pamcard/Envio/updateVlrParcelaContrato_{0}_{1}.xml", DateTime.Now().ToString("yyyy-MM-dd"), Contrato.Codigo))
            Dim xmlArquivo As New XmlDocument()
            xmlArquivo.LoadXml(strXML)
            xmlArquivo.Save(strNomeArquivo)
            xmlArquivo = Nothing

            Dim wsFields(ds.Tables(0).Rows.Count - 1) As WSPamcard.FieldTO
            For i = 0 To ds.Tables(0).Rows.Count - 1
                Dim teste As String = i
                SetValues(ds.Tables(0).Rows(i).Item("Codigo"), ds.Tables(0).Rows(i).Item("Valor"), wsFields, i)
            Next

            'Dim wsRequest As New WsPamcard.RequestTO()
            'wsRequest.certificate = X509Certificate.CreateFromCertFile(HttpContext.Current.Server.MapPath("Pamcard/pamcard_alvorada_04854422000185.crt")).GetRawCertData()
            'wsRequest.context = "UpdateValuesFreightContract"
            'wsRequest.fields = wsFields

            'Dim wsPamCard As New WsPamcard.WSPamcardService()
            'Dim wsResponse As New WsPamcard.ResponseTO()
            'wsResponse = wsPamCard.execute(wsRequest)

            'Dim wsFieldsResponse() As WsPamcard.FieldTO = wsResponse.fields

            'strXML = "<updateValuesFreightContract>"
            'strXML &= String.Format("<erro>{0}</erro><errodescricao>{1}</errodescricao>", _
            '                        GetValue("mensagem.codigo", wsFieldsResponse), _
            '                        GetValue("mensagem.descricao", wsFieldsResponse))
            'strXML &= "</updateValuesFreightContract>"

            strXML = "<updateValuesFreightContract>"
            strXML &= String.Format("<erro>{0}</erro><errodescricao>{1}</errodescricao>", _
                                    "0", _
                                    "TUDO CERTO")
            strXML &= "</updateValuesFreightContract>"

            strNomeArquivo = HttpContext.Current.Server.MapPath(String.Format("Pamcard/Retorno/respUpdateVlrParcelaContrato_{0}_{1}.xml", DateTime.Now().ToString("yyyy-MM-dd"), Contrato.Codigo))

            xmlArquivo = New XmlDocument()
            xmlArquivo.LoadXml(strXML)
            xmlArquivo.Save(strNomeArquivo)
            xmlArquivo = Nothing

            dsRet.ReadXml(strNomeArquivo)

            Return dsRet
        Catch ex As Exception
            HttpContext.Current.Session("ssMessage") = Funcoes.EliminarCaracteresEspeciais(ex.Message.ToString)
            dsRet = Nothing
            Return dsRet
        End Try
    End Function

    Public Function IncluirFavorecido(ByVal Empresa As String, ByVal DDD As String, ByVal Fone As String, ByVal Favorecido As Cliente) As DataSet
        Dim ds As New DataSet
        Dim dsRet As New DataSet

        Try
            Dim tbP As New DataTable("Pamcard")
            tbP.Columns.Add("Codigo", Type.GetType("System.String"))
            tbP.Columns.Add("Valor", Type.GetType("System.String"))
            ds.Tables.Add(tbP)

            SetDataset("viagem.contratante.documento.numero", Empresa, ds)

            SetDataset("viagem.favorecido.documento.qtde", 2, ds)
            If Favorecido.Codigo.Length = 11 Then
                SetDataset("viagem.favorecido.documento1.tipo", 2, ds)
            Else
                SetDataset("viagem.favorecido.documento1.tipo", 1, ds)
            End If
            SetDataset("viagem.favorecido.documento1.numero", Favorecido.Codigo, ds)
            If Favorecido.Codigo.Length = 11 Then
                SetDataset("viagem.favorecido.documento2.tipo", 5, ds)
            Else
                SetDataset("viagem.favorecido.documento2.tipo", 6, ds)
            End If
            SetDataset("viagem.favorecido.documento2.numero", Favorecido.RNTRCTransportador, ds)

            SetDataset("viagem.favorecido.nome", Favorecido.Nome, ds)
            SetDataset("viagem.favorecido.endereco.logradouro", Favorecido.Endereco, ds)
            SetDataset("viagem.favorecido.endereco.numero", Favorecido.Numero, ds)
            SetDataset("viagem.favorecido.endereco.complemento", Favorecido.Complemento, ds)
            SetDataset("viagem.favorecido.endereco.bairro", Favorecido.Bairro, ds)
            SetDataset("viagem.favorecido.endereco.cidade.ibge", Favorecido.Municipio.EstadoIbge & Favorecido.CodigoMunicipio.ToString("00000"), ds)
            SetDataset("viagem.favorecido.endereco.cep", Funcoes.OnlyNumbers(Favorecido.CEP), ds)

            SetDataset("viagem.favorecido.telefone.ddd", DDD, ds)
            SetDataset("viagem.favorecido.telefone.numero", Fone, ds)

            Dim strXML As String

            strXML = "<insertFavorecido>"
            For Each dr As DataRow In ds.Tables(0).Rows
                strXML &= "<" & dr("Codigo") & ">" & dr("Valor") & "</" & dr("Codigo") & ">"
            Next
            strXML &= "</insertFavorecido>"

            Dim strNomeArquivo As String = HttpContext.Current.Server.MapPath(String.Format("Pamcard/Envio/insertFavorecido_{0}_{1}.xml", DateTime.Now().ToString("yyyy-MM-dd"), Favorecido.Codigo))
            Dim xmlArquivo As New XmlDocument()
            xmlArquivo.LoadXml(strXML)
            xmlArquivo.Save(strNomeArquivo)
            xmlArquivo = Nothing

            Dim wsFields(ds.Tables(0).Rows.Count - 1) As WSPamcard.FieldTO
            Dim i As Integer
            For i = 0 To ds.Tables(0).Rows.Count - 1
                Dim teste As String = i
                SetValues(ds.Tables(0).Rows(i).Item("Codigo"), ds.Tables(0).Rows(i).Item("Valor"), wsFields, i)
            Next

            Dim wsRequest As New WSPamcard.RequestTO()
            wsRequest.certificate = X509Certificate.CreateFromCertFile(HttpContext.Current.Server.MapPath("Pamcard/pamcard_alvorada_04854422000185.crt")).GetRawCertData()
            wsRequest.context = "InsertFavored"
            wsRequest.fields = wsFields
            '--------------------------------------------------------------------------------------------

            Dim wsPamCard As New WSPamcard.WSPamcardService()
            Dim wsResponse As New WSPamcard.ResponseTO()
            wsResponse = wsPamCard.execute(wsRequest)

            Dim wsFieldsResponse() As WSPamcard.FieldTO = wsResponse.fields

            strXML = "<insertFavorecido>"
            strXML &= String.Format("<erro>{0}</erro><errodescricao>{1}</errodescricao>", _
                                    GetValue("mensagem.codigo", wsFieldsResponse), _
                                    GetValue("mensagem.descricao", wsFieldsResponse))
            strXML &= "</insertFavorecido>"

            strNomeArquivo = HttpContext.Current.Server.MapPath(String.Format("Pamcard/Retorno/respInsertFavorecido_{0}_{1}.xml", DateTime.Now().ToString("yyyy-MM-dd"), Favorecido.Codigo))

            xmlArquivo = New XmlDocument()
            xmlArquivo.LoadXml(strXML)
            xmlArquivo.Save(strNomeArquivo)
            xmlArquivo = Nothing

            Dim dsr As New DataSet
            dsr.ReadXml(strNomeArquivo)

            Return dsr
        Catch ex As Exception
            HttpContext.Current.Session("ssMessage") = Funcoes.EliminarCaracteresEspeciais(ex.Message.ToString)
            dsRet = Nothing
            Return dsRet
        End Try

    End Function

    'Public Function ConsultarContratoDeFrete() As String
    '    Dim wsFields(12) As WSPamcard.FieldTO

    '    SetValues("viagem.id", 1035601, wsFields, 0)
    '    SetValues("viagem.id.cliente", "", wsFields, 1)
    '    SetValues("viagem.antt.ciot.numero", "", wsFields, 2)

    '    SetValues("viagem.contratante.documento.numero", "27064195000190", wsFields, 3)
    '    SetValues("viagem.unidade.documento.tipo", 1, wsFields, 4)
    '    SetValues("viagem.unidade.documento.numero", "27064195000270", wsFields, 5)

    '    SetValues("viagem.pedagio.obter.praca", "N", wsFields, 6)
    '    SetValues("viagem.pedagio.obter.rota", "N", wsFields, 7)
    '    SetValues("viagem.obter.favorecido", "N", wsFields, 8)
    '    SetValues("viagem.obter.documento", "N", wsFields, 9)
    '    SetValues("viagem.obter.valores", "N", wsFields, 10)
    '    SetValues("viagem.obter.veiculo", "N", wsFields, 11)
    '    SetValues("viagem.obter.quitacao", "N", wsFields, 12)

    '    Dim wsRequest As New WSPamcard.RequestTO()
    '    '--------------------------------------------------------------------------------------------

    '    Return ""

    'End Function

    'Public Function ConsultarStatusDaParcela() As String
    '    Dim wsFields(7) As WSPamcard.FieldTO

    '    SetValues("viagem.id", 1026224, wsFields, 0)
    '    SetValues("viagem.id.cliente", "", wsFields, 1)
    '    SetValues("viagem.antt.ciot.numero", "", wsFields, 2)

    '    SetValues("viagem.contratante.documento.numero", "27064195000190", wsFields, 3)
    '    SetValues("viagem.unidade.documento.tipo", 1, wsFields, 4)
    '    SetValues("viagem.unidade.documento.numero", "27064195000270", wsFields, 5)

    '    SetValues("viagem.parcela.numero", 6, wsFields, 6)
    '    SetValues("viagem.parcela.numero.cliente", 10, wsFields, 7)


    '    Dim wsRequest As New WSPamcard.RequestTO()
    '    '--------------------------------------------------------------------------------------------

    '    Return ""

    'End Function

    'Public Function AlterarStatusDaParcela() As String
    '    Dim wsFields(9) As WSPamcard.FieldTO

    '    SetValues("viagem.id", 1026224, wsFields, 0)
    '    SetValues("viagem.id.cliente", "", wsFields, 1)
    '    SetValues("viagem.antt.ciot.numero", "", wsFields, 2)

    '    SetValues("viagem.contratante.documento.numero", "27064195000190", wsFields, 3)
    '    SetValues("viagem.unidade.documento.tipo", 1, wsFields, 4)
    '    SetValues("viagem.unidade.documento.numero", "27064195000270", wsFields, 5)

    '    SetValues("viagem.parcela.qtde", 1, wsFields, 6)
    '    SetValues("viagem.parcela1.numero", 3, wsFields, 7)
    '    SetValues("viagem.parcela1.numero.cliente", 10, wsFields, 8)
    '    SetValues("viagem.parcela1.status.id", 3, wsFields, 9)


    '    Dim wsRequest As New WSPamcard.RequestTO()
    '    '--------------------------------------------------------------------------------------------

    '    Return ""

    'End Function

    'Public Function PagamentoDeParcela() As String
    '    Dim wsFields(8) As WSPamcard.FieldTO

    '    SetValues("viagem.id", 1035589, wsFields, 0)
    '    SetValues("viagem.id.cliente", "", wsFields, 1)
    '    SetValues("viagem.antt.ciot.numero", "", wsFields, 2)

    '    SetValues("viagem.contratante.documento.numero", "27064195000190", wsFields, 3)
    '    SetValues("viagem.unidade.documento.tipo", 1, wsFields, 4)
    '    SetValues("viagem.unidade.documento.numero", "27064195000270", wsFields, 5)

    '    SetValues("viagem.parcela.qtde", 1, wsFields, 6)

    '    SetValues("viagem.parcela1.numero", 2, wsFields, 7)
    '    SetValues("viagem.parcela1.numero.cliente", 10, wsFields, 8)


    '    Dim wsRequest As New WSPamcard.RequestTO()
    '    '--------------------------------------------------------------------------------------------

    '    Return ""

    'End Function

    'Public Function ConsultaStatusDoPedagio() As String
    '    Dim wsFields(5) As WSPamcard.FieldTO

    '    SetValues("viagem.contratante.documento.numero", "27064195000190", wsFields, 0)
    '    SetValues("viagem.unidade.documento.tipo", 1, wsFields, 1)
    '    SetValues("viagem.unidade.documento.numero", "27064195000270", wsFields, 2)

    '    SetValues("viagem.id", 1035589, wsFields, 3)
    '    SetValues("viagem.id.cliente", "", wsFields, 4)
    '    SetValues("viagem.antt.ciot.numero", "", wsFields, 5)

    '    Dim wsRequest As New WSPamcard.RequestTO()
    '    '--------------------------------------------------------------------------------------------

    '    Return ""

    'End Function

    'Public Function AlterarStatusDoPedagio() As String
    '    Dim wsFields(6) As WSPamcard.FieldTO

    '    SetValues("viagem.contratante.documento.numero", "27064195000190", wsFields, 0)
    '    SetValues("viagem.unidade.documento.tipo", 1, wsFields, 1)
    '    SetValues("viagem.unidade.documento.numero", "27064195000270", wsFields, 2)

    '    SetValues("viagem.id", 1026229, wsFields, 3)
    '    SetValues("viagem.id.cliente", "", wsFields, 4)
    '    SetValues("viagem.antt.ciot.numero", "", wsFields, 5)

    '    SetValues("viagem.pedagio.status.id", 3, wsFields, 6)

    '    Dim wsRequest As New WSPamcard.RequestTO()
    '    '--------------------------------------------------------------------------------------------

    '    Return ""

    'End Function

    'Public Function Roteirizar() As String
    '    Dim wsFields(29) As WSPamcard.FieldTO

    '    SetValues("viagem.contratante.documento.numero", "27064195000190", wsFields, 0)
    '    SetValues("viagem.unidade.documento.tipo", 1, wsFields, 1)
    '    SetValues("viagem.unidade.documento.numero", "27064195000270", wsFields, 2)

    '    SetValues("viagem.veiculo.categoria", 7, wsFields, 3)

    '    SetValues("viagem.rota.id", 1234543, wsFields, 4)

    '    SetValues("viagem.rota.nome", "ROTA-SPO-CMP", wsFields, 5)

    '    SetValues("viagem.origem.pais.nome", "BRASIL", wsFields, 6)
    '    SetValues("viagem.origem.estado.nome", "SP", wsFields, 7)
    '    SetValues("viagem.origem.cidade.nome", "SAO PAULO", wsFields, 8)

    '    SetValues("viagem.ponto.qtde", 3, wsFields, 9)

    '    SetValues("viagem.ponto1.pais.nome", "BRASIL", wsFields, 10)
    '    SetValues("viagem.ponto1.estado.nome", "SP", wsFields, 11)
    '    SetValues("viagem.ponto1.cidade.nome", "MAIRIPORA", wsFields, 12)

    '    SetValues("viagem.ponto2.pais.nome", "BRASIL", wsFields, 13)
    '    SetValues("viagem.ponto2.estado.nome", "SP", wsFields, 14)
    '    SetValues("viagem.ponto2.cidade.nome", "JUNDIAI", wsFields, 15)

    '    SetValues("viagem.ponto3.pais.nome", "BRASIL", wsFields, 16)
    '    SetValues("viagem.ponto3.estado.nome", "SP", wsFields, 17)
    '    SetValues("viagem.ponto3.cidade.nome", "CAMPINAS", wsFields, 18)

    '    SetValues("viagem.destino.pais.nome", "BRASIL", wsFields, 19)
    '    SetValues("viagem.destino.estado.nome", "SP", wsFields, 20)
    '    SetValues("viagem.destino.cidade.nome", "CAMPINAS", wsFields, 21)

    '    SetValues("viagem.pedagio.obter.rota", "S", wsFields, 22)

    '    SetValues("viagem.origem.cidade.ibge", "4205456", wsFields, 23)
    '    SetValues("viagem.destino.cidade.ibge", "4208203", wsFields, 24)
    '    SetValues("viagem.ponto.qtde", 2, wsFields, 25)
    '    SetValues("viagem.ponto1.cidade.ibge", "4218707", wsFields, 26)
    '    SetValues("viagem.ponto2.cidade.ibge", "4211900", wsFields, 27)

    '    SetValues("viagem.rota.id", 10395, wsFields, 28)

    '    SetValues("viagem.rota.nome", "sao paulo - campinas", wsFields, 29)



    '    Dim wsRequest As New WSPamcard.RequestTO()
    '    '--------------------------------------------------------------------------------------------

    '    Return ""

    'End Function

    'Public Function IncluirCartaoPortadorFrete() As String
    '    Dim wsFields(21) As WSPamcard.FieldTO

    '    SetValues("viagem.contratante.documento.numero", "27064195000190", wsFields, 0)

    '    SetValues("viagem.cartao.numero", 4417810094174019, wsFields, 1)

    '    SetValues("viagem.cartao.portador.documento.numero", 60913811211, wsFields, 2)
    '    SetValues("viagem.cartao.portador.documento.tipo", 1, wsFields, 3)

    '    SetValues("viagem.cartao.portador.rg", 321654, wsFields, 4)
    '    SetValues("viagem.cartao.portador.uf.rg", "SP", wsFields, 5)
    '    SetValues("viagem.cartao.portador.nome", "jennifer teste", wsFields, 6)
    '    SetValues("viagem.cartao.portador.rntrc", 123456879, wsFields, 7)
    '    SetValues("viagem.cartao.portador.data.nascimento", "28/12/2981", wsFields, 8)
    '    SetValues("viagem.cartao.portador.endereco.logradouro", "R. LOURENCO CABRAL", wsFields, 9)
    '    SetValues("viagem.cartao.portador.endereco.numero", "843", wsFields, 10)
    '    SetValues("viagem.cartao.portador.endereco.complemento", "COMPLEMENTO", wsFields, 11)
    '    SetValues("viagem.cartao.portador.endereco.bairro", "VILA GAGA", wsFields, 12)
    '    SetValues("viagem.cartao.portador.endereco.cidade", "SAO PAULO", wsFields, 13)
    '    SetValues("viagem.cartao.portador.endereco.uf", "SP", wsFields, 14)
    '    SetValues("viagem.cartao.portador.endereco.pais", "BRASIL", wsFields, 15)
    '    SetValues("viagem.cartao.portador.endereco.cep", "04152130", wsFields, 16)
    '    SetValues("viagem.cartao.portador.telefone.ddd", "011", wsFields, 17)
    '    SetValues("viagem.cartao.portador.telefone.numero", "38891105", wsFields, 18)
    '    SetValues("viagem.cartao.portador.celular.ddd", "013", wsFields, 19)
    '    SetValues("viagem.cartao.portador.celular.numero", "99995555", wsFields, 20)
    '    SetValues("viagem.cartao.portador.email", "jennifer.cavalheiro@gps-pamcary.com.br", wsFields, 21)


    '    Dim wsRequest As New WSPamcard.RequestTO()
    '    '--------------------------------------------------------------------------------------------

    '    Return ""

    'End Function


    'Public Function IncluirConta() As String
    '    Dim wsFields(6) As WSPamcard.FieldTO

    '    SetValues("viagem.contratante.documento.numero", "00001001000312", wsFields, 0)

    '    SetValues("viagem.favorecido.documento.tipo", 2, wsFields, 1)
    '    SetValues("viagem.favorecido.documento.numero", "38666615338", wsFields, 2)

    '    SetValues("viagem.favorecido.conta.banco", 345, wsFields, 3)
    '    SetValues("viagem.favorecido.conta.agencia", "1234", wsFields, 4)
    '    SetValues("viagem.favorecido.conta.numero", "987654323", wsFields, 5)
    '    SetValues("viagem.favorecido.conta.tipo", "2", wsFields, 6)



    '    Dim wsRequest As New WSPamcard.RequestTO()
    '    '--------------------------------------------------------------------------------------------

    '    Return ""

    'End Function

    'Public Function ConsultarConta() As String
    '    Dim wsFields(6) As WSPamcard.FieldTO

    '    SetValues("viagem.contratante.documento.numero", "00001001000312", wsFields, 0)

    '    SetValues("viagem.favorecido.documento.tipo", 2, wsFields, 1)
    '    SetValues("viagem.favorecido.documento.numero", "38666615338", wsFields, 2)

    '    SetValues("viagem.favorecido.conta.banco", 345, wsFields, 3)
    '    SetValues("viagem.favorecido.conta.agencia", "1234", wsFields, 4)
    '    SetValues("viagem.favorecido.conta.numero", "987654323", wsFields, 5)
    '    SetValues("viagem.favorecido.conta.tipo", "1", wsFields, 6)



    '    Dim wsRequest As New WSPamcard.RequestTO()
    '    '--------------------------------------------------------------------------------------------

    '    Return ""

    'End Function

    'Public Function CancelarContratoDeFrete() As String
    '    Dim wsFields(6) As WSPamcard.FieldTO

    '    SetValues("viagem.contratante.documento.numero", "27064195000190", wsFields, 0)
    '    SetValues("viagem.unidade.documento.tipo", 1, wsFields, 1)
    '    SetValues("viagem.unidade.documento.numero", "27064195000270", wsFields, 2)

    '    SetValues("viagem.id", 1035598, wsFields, 3)
    '    SetValues("viagem.id.cliente", "", wsFields, 4)
    '    SetValues("viagem.antt.ciot.numero", "", wsFields, 5)
    '    SetValues("viagem.antt.cancelamento.motivo", "", wsFields, 6)



    '    Dim wsRequest As New WSPamcard.RequestTO()
    '    '--------------------------------------------------------------------------------------------

    '    Return ""

    'End Function

    Public Function ConsultarCartao(ByVal Empresa As String, ByVal Cartao As String) As DataSet
        Dim wsFields(1) As WSPamcard.FieldTO

        SetValues("viagem.contratante.documento.numero", Empresa, wsFields, 0)
        SetValues("viagem.cartao.numero", Cartao, wsFields, 1)

        Dim wsRequest As New WSPamcard.RequestTO()
        wsRequest.certificate = X509Certificate.CreateFromCertFile(HttpContext.Current.Server.MapPath("Pamcard/pamcard_alvorada_04854422000185.crt")).GetRawCertData()
        wsRequest.context = "FindCard"
        wsRequest.fields = wsFields

        Dim wsPamCard As New WSPamcard.WSPamcardService()
        Dim strXML As String

        strXML = "<rntrcCartao>"
        strXML &= String.Format("<cnpjcontr>{0}</cnpjcontr><cartao>{1}</cartao>", _
                                GetValue("viagem.contratante.documento.numero", wsFields), _
                                GetValue("viagem.cartao.numero", wsFields))
        strXML &= "</rntrcCartao>"

        ' TransacoesPamcard
        Dim strNomeArquivo As String = HttpContext.Current.Server.MapPath(String.Format("Pamcard/Envio/findCard_{0}_{1}.xml", DateTime.Now().ToString("yyyy-MM-dd-HH-mm-ss"), Cartao))

        Dim xmlArquivo As New XmlDocument()
        xmlArquivo.LoadXml(strXML)
        xmlArquivo.Save(strNomeArquivo)
        xmlArquivo = Nothing

        Dim wsResponse As New WSPamcard.ResponseTO()
        wsResponse = wsPamCard.execute(wsRequest)

        Dim wsFieldsResponse() As WSPamcard.FieldTO = wsResponse.fields

        strXML = "<rntrcCartao>"
        strXML &= String.Format("<erro>{0}</erro><errodescricao>{1}</errodescricao><numero>{2}</numero><tipo>{3}</tipo><nome>{4}</nome><descricao>{5}</descricao><id>{6}</id>", _
                                GetValue("mensagem.codigo", wsFieldsResponse), _
                                GetValue("mensagem.descricao", wsFieldsResponse), _
                                GetValue("viagem.cartao.portador.documento.numero", wsFieldsResponse), _
                                GetValue("viagem.cartao.portador.documento.tipo", wsFieldsResponse), _
                                GetValue("viagem.cartao.portador.nome", wsFieldsResponse), _
                                GetValue("viagem.cartao.status.descricao", wsFieldsResponse), _
                                GetValue("viagem.cartao.status.id", wsFieldsResponse))
        strXML &= "</rntrcCartao>"

        strNomeArquivo = HttpContext.Current.Server.MapPath(String.Format("Pamcard/Retorno/respCard_{0}_{1}.xml", DateTime.Now().ToString("yyyy-MM-dd-HH-mm-ss"), Cartao))

        xmlArquivo = New XmlDocument()
        xmlArquivo.LoadXml(strXML)
        xmlArquivo.Save(strNomeArquivo)
        xmlArquivo = Nothing

        Dim ds As New DataSet
        ds.ReadXml(strNomeArquivo)

        Return ds
    End Function

    Public Function ConsultarFavorecido(ByVal Empresa As String, ByVal Favorecido As String) As DataSet
        Dim wsFields(2) As WSPamcard.FieldTO

        SetValues("viagem.contratante.documento.numero", Empresa, wsFields, 0)
        If Favorecido.Length = 11 Then
            SetValues("viagem.favorecido.documento.tipo", 2, wsFields, 1)
        Else
            SetValues("viagem.favorecido.documento.tipo", 1, wsFields, 1)
        End If
        SetValues("viagem.favorecido.documento.numero", Favorecido, wsFields, 2)

        Dim wsRequest As New WSPamcard.RequestTO()
        wsRequest.certificate = X509Certificate.CreateFromCertFile(HttpContext.Current.Server.MapPath("Pamcard/pamcard_alvorada_04854422000185.crt")).GetRawCertData()
        wsRequest.context = "FindFavored"
        wsRequest.fields = wsFields

        Dim wsPamCard As New WSPamcard.WSPamcardService()
        Dim strXML As String

        strXML = "<rntrcFavorecido>"
        strXML &= String.Format("<cnpjcontr>{0}</cnpjcontr><favorecido>{1}</favorecido>", _
                                GetValue("viagem.contratante.documento.numero", wsFields), _
                                GetValue("viagem.favorecido.documento.numero", wsFields))
        strXML &= "</rntrcFavorecido>"

        ' TransacoesPamcard
        Dim strNomeArquivo As String = HttpContext.Current.Server.MapPath(String.Format("Pamcard/Envio/findFavorecido_{0}_{1}.xml", DateTime.Now().ToString("yyyy-MM-dd-HH-mm-ss"), Favorecido))

        Dim xmlArquivo As New XmlDocument()
        xmlArquivo.LoadXml(strXML)
        xmlArquivo.Save(strNomeArquivo)
        xmlArquivo = Nothing

        Dim wsResponse As New WSPamcard.ResponseTO()
        wsResponse = wsPamCard.execute(wsRequest)

        Dim wsFieldsResponse() As WSPamcard.FieldTO = wsResponse.fields

        Dim intCartoes As Integer = GetValue("viagem.favorecido.cartao.qtde", wsFieldsResponse)
        Dim intContas As Integer = GetValue("viagem.favorecido.conta.qtde", wsFieldsResponse)

        strXML = "<rntrcFavorecido>"
        strXML &= String.Format("<erro>{0}</erro><errodescricao>{1}</errodescricao><qtdeCartao>{2}</qtdeCartao>", _
                                GetValue("mensagem.codigo", wsFieldsResponse), _
                                GetValue("mensagem.descricao", wsFieldsResponse), _
                                intCartoes)

        For intCartao As Integer = 1 To intCartoes
            strXML &= "<cartao>"
            strXML &= String.Format("<cNumero>{0}</cNumero><cTipo>{1}</cTipo><cStatus>{2}</cStatus>", _
                                    GetValue("viagem.favorecido.cartao" & intCartao & ".numero", wsFieldsResponse), _
                                    GetValue("viagem.favorecido.cartao" & intCartao & ".tipo", wsFieldsResponse), _
                                    GetValue("viagem.favorecido.cartao" & intCartao & ".status", wsFieldsResponse))
            strXML &= "</cartao>"
        Next

        strXML &= "<ctaQdte>" & intContas & "</ctaQdte>"

        For intConta As Integer = 1 To intContas
            strXML &= "<conta>"
            strXML &= String.Format("<ctaBanco>{0}</ctaBanco><ctaAgencia>{1}</ctaAgencia><ctaNumero>{2}</ctaNumero><ctaTipo>{3}</ctaTipo><ctaStatus>{4}</ctaStatus>", _
                                    GetValue("viagem.favorecido.conta" & intConta & ".banco", wsFieldsResponse), _
                                    GetValue("viagem.favorecido.conta" & intConta & ".agencia", wsFieldsResponse), _
                                    GetValue("viagem.favorecido.conta" & intConta & ".numero", wsFieldsResponse), _
                                    GetValue("viagem.favorecido.conta" & intConta & ".tipo", wsFieldsResponse), _
                                    GetValue("viagem.favorecido.conta" & intConta & ".status", wsFieldsResponse))
            strXML &= "</conta>"
        Next

        strXML &= String.Format("<favNome>{0}</favNome><favStatus>{1}</favStatus>", _
                                GetValue("viagem.favorecido.nome", wsFieldsResponse), _
                                GetValue("viagem.favorecido.status.rntrc", wsFieldsResponse))

        strXML &= "</rntrcFavorecido>"

        strNomeArquivo = HttpContext.Current.Server.MapPath(String.Format("Pamcard/Retorno/respFavorecido_{0}_{1}.xml", DateTime.Now().ToString("yyyy-MM-dd-HH-mm-ss"), Favorecido))

        xmlArquivo = New XmlDocument()
        xmlArquivo.LoadXml(strXML)
        xmlArquivo.Save(strNomeArquivo)
        xmlArquivo = Nothing

        Dim ds As New DataSet
        ds.ReadXml(strNomeArquivo)

        Return ds
    End Function

    Public Function ConsultarRNTRC(ByVal Empresa As String, ByVal Transportador As String, ByVal codigoRNTRC As String) As DataSet
        Dim ds As New DataSet
        Dim dsRet As New DataSet

        Try
            Dim tbP As New DataTable("Pamcard")
            tbP.Columns.Add("Codigo", Type.GetType("System.String"))
            tbP.Columns.Add("Valor", Type.GetType("System.String"))
            ds.Tables.Add(tbP)

            SetDataset("viagem.contratante.documento.numero", Empresa, ds)

            SetDataset("viagem.favorecido.documento.qtde", 2, ds)
            If Transportador.Length = 11 Then
                SetDataset("viagem.favorecido.documento1.tipo", 2, ds)
            Else
                SetDataset("viagem.favorecido.documento1.tipo", 1, ds)
            End If
            SetDataset("viagem.favorecido.documento1.numero", Transportador, ds)
            If Transportador.Length = 11 Then
                SetDataset("viagem.favorecido.documento2.tipo", 5, ds)
            Else
                SetDataset("viagem.favorecido.documento2.tipo", 6, ds)
            End If
            SetDataset("viagem.favorecido.documento2.numero", codigoRNTRC, ds)

            Dim strXML As String
            strXML = "<findRNTRC>"
            For Each dr As DataRow In ds.Tables(0).Rows
                strXML &= "<" & dr("Codigo") & ">" & dr("Valor") & "</" & dr("Codigo") & ">"
            Next
            strXML &= "</findRNTRC>"

            Dim strNomeArquivo As String = HttpContext.Current.Server.MapPath(String.Format("Pamcard/Envio/findRNTRC_{0}_{1}.xml", DateTime.Now().ToString("yyyy-MM-dd-HH-mm-ss"), Transportador))

            Dim xmlArquivo As New XmlDocument()
            xmlArquivo.LoadXml(strXML)
            xmlArquivo.Save(strNomeArquivo)
            xmlArquivo = Nothing

            Dim wsFields(ds.Tables(0).Rows.Count - 1) As WSPamcard.FieldTO
            Dim i As Integer
            For i = 0 To ds.Tables(0).Rows.Count - 1
                Dim teste As String = i
                SetValues(ds.Tables(0).Rows(i).Item("Codigo"), ds.Tables(0).Rows(i).Item("Valor"), wsFields, i)
            Next

            Dim wsRequest As New WSPamcard.RequestTO()
            wsRequest.certificate = X509Certificate.CreateFromCertFile(HttpContext.Current.Server.MapPath("Pamcard/pamcard_alvorada_04854422000185.crt")).GetRawCertData()
            wsRequest.context = "FindRNTRC"
            wsRequest.fields = wsFields

            Dim wsPamCard As New WSPamcard.WSPamcardService()
            Dim wsResponse As New WSPamcard.ResponseTO()
            wsResponse = wsPamCard.execute(wsRequest)

            Dim wsFieldsResponse() As WSPamcard.FieldTO = wsResponse.fields

            If wsFieldsResponse.Length = 2 Then
                strXML = "<respRNTRC>"
                strXML &= String.Format("<erro>{0}</erro><errodescricao>{1}</errodescricao>", _
                                        GetValue("mensagem.codigo", wsFieldsResponse), _
                                        GetValue("mensagem.descricao", wsFieldsResponse))
                strXML &= "</respRNTRC>"
            Else
                strXML = "<respRNTRC>"
                strXML &= String.Format("<erro>{0}</erro><errodescricao>{1}</errodescricao><nome>{2}</nome><tac>{3}</tac><situacao>{4}</situacao><tipo>{5}</tipo><validade>{6}</validade>", _
                                        GetValue("mensagem.codigo", wsFieldsResponse), _
                                        GetValue("mensagem.descricao", wsFieldsResponse), _
                                        GetValue("viagem.antt.nome", wsFieldsResponse).ToString.Replace("&", "E"), _
                                        GetValue("viagem.antt.rntrc.equiparado.tac", wsFieldsResponse), _
                                        GetValue("viagem.antt.rntrc.situacao", wsFieldsResponse), _
                                        GetValue("viagem.antt.rntrc.tipo", wsFieldsResponse), _
                                        GetValue("viagem.antt.rntrc.validade", wsFieldsResponse))
                strXML &= "</respRNTRC>"
            End If


            strNomeArquivo = HttpContext.Current.Server.MapPath(String.Format("Pamcard/Retorno/respRNTRC_{0}_{1}.xml", DateTime.Now().ToString("yyyy-MM-dd-HH-mm-ss"), Transportador))

            xmlArquivo = New XmlDocument()
            xmlArquivo.LoadXml(strXML)
            xmlArquivo.Save(strNomeArquivo)
            xmlArquivo = Nothing

            dsRet.ReadXml(strNomeArquivo)

            Return dsRet
        Catch ex As Exception
            HttpContext.Current.Session("ssMessage") = Funcoes.EliminarCaracteresEspeciais(ex.Message.ToString)
            Return dsRet
        End Try
    End Function

    Public Function ConsultarFrota(ByVal Empresa As String, ByVal Transportador As String, ByVal codigoRNTRC As String, ByVal Placa As Placa) As DataSet
        Dim ds As New DataSet

        Try
            Dim i As Integer = 7
            Dim j As Integer = 1

            If Placa.Placa02.Length > 0 Then
                i += 1
                j += 1
            End If

            If Placa.Placa03.Length > 0 Then
                i += 1
                j += 1
            End If

            If Placa.Placa04.Length > 0 Then
                i += 1
                j += 1
            End If

            Dim wsFields(i) As WSPamcard.FieldTO

            SetValues("viagem.contratante.documento.numero", Empresa, wsFields, 0)
            SetValues("viagem.favorecido.documento.qtde", 2, wsFields, 1)
            If Transportador.Length = 11 Then
                SetValues("viagem.favorecido.documento1.tipo", 2, wsFields, 2)
            Else
                SetValues("viagem.favorecido.documento1.tipo", 1, wsFields, 2)
            End If
            SetValues("viagem.favorecido.documento1.numero", Transportador, wsFields, 3)
            If Transportador.Length = 11 Then
                SetValues("viagem.favorecido.documento2.tipo", 5, wsFields, 4)
            Else
                SetValues("viagem.favorecido.documento2.tipo", 6, wsFields, 4)
            End If
            SetValues("viagem.favorecido.documento2.numero", codigoRNTRC, wsFields, 5)

            SetValues("viagem.veiculo.qtde", j, wsFields, 6)

            SetValues("viagem.veiculo1.placa", Placa.Placa01.Replace("-", ""), wsFields, 7)

            If Placa.Placa02.Length > 0 Then
                SetValues("viagem.veiculo2.placa", Placa.Placa02.Replace("-", ""), wsFields, 8)
            End If

            If Placa.Placa03.Length > 0 Then
                SetValues("viagem.veiculo3.placa", Placa.Placa03.Replace("-", ""), wsFields, 9)
            End If

            If Placa.Placa04.Length > 0 Then
                SetValues("viagem.veiculo4.placa", Placa.Placa04.Replace("-", ""), wsFields, 10)
            End If

            Dim wsRequest As New WSPamcard.RequestTO()
            wsRequest.certificate = X509Certificate.CreateFromCertFile(HttpContext.Current.Server.MapPath("Pamcard/pamcard_alvorada_04854422000185.crt")).GetRawCertData()
            wsRequest.context = "FindFleet"
            wsRequest.fields = wsFields

            Dim wsPamCard As New WSPamcard.WSPamcardService()
            Dim strXML As String

            strXML = "<frota>"
            Dim strFrota As String = "<cnpjcontr>{0}</cnpjcontr><qtde>{1}</qtde><tipo>{2}</tipo><numero>{3}</numero><tipo>{4}</tipo><numero>{5}</numero><qtde>{6}</qtde><placa>{7}</placa>"

            If Placa.Placa04.Length > 0 Then
                strFrota &= "<placa>{8}</placa>"
                strFrota &= "<placa>{9}</placa>"
                strFrota &= "<placa>{10}</placa>"
                strXML &= String.Format(strFrota, _
                                    GetValue("viagem.contratante.documento.numero", wsFields), _
                                    GetValue("viagem.favorecido.documento.qtde", wsFields), _
                                    GetValue("viagem.favorecido.documento1.tipo", wsFields), _
                                    GetValue("viagem.favorecido.documento1.numero", wsFields), _
                                    GetValue("viagem.favorecido.documento2.tipo", wsFields), _
                                    GetValue("viagem.favorecido.documento2.numero", wsFields), _
                                    GetValue("viagem.veiculo.qtde", wsFields), _
                                    GetValue("viagem.veiculo1.placa", wsFields), _
                                    GetValue("viagem.veiculo2.placa", wsFields), _
                                    GetValue("viagem.veiculo3.placa", wsFields), _
                                    GetValue("viagem.veiculo4.placa", wsFields))
            ElseIf Placa.Placa03.Length > 0 Then
                strFrota &= "<placa>{8}</placa>"
                strFrota &= "<placa>{9}</placa>"
                strXML &= String.Format(strFrota, _
                                    GetValue("viagem.contratante.documento.numero", wsFields), _
                                    GetValue("viagem.favorecido.documento.qtde", wsFields), _
                                    GetValue("viagem.favorecido.documento1.tipo", wsFields), _
                                    GetValue("viagem.favorecido.documento1.numero", wsFields), _
                                    GetValue("viagem.favorecido.documento2.tipo", wsFields), _
                                    GetValue("viagem.favorecido.documento2.numero", wsFields), _
                                    GetValue("viagem.veiculo.qtde", wsFields), _
                                    GetValue("viagem.veiculo1.placa", wsFields), _
                                    GetValue("viagem.veiculo2.placa", wsFields), _
                                    GetValue("viagem.veiculo3.placa", wsFields))
            ElseIf Placa.Placa02.Length > 0 Then
                strFrota &= "<placa>{8}</placa>"
                strXML &= String.Format(strFrota, _
                                    GetValue("viagem.contratante.documento.numero", wsFields), _
                                    GetValue("viagem.favorecido.documento.qtde", wsFields), _
                                    GetValue("viagem.favorecido.documento1.tipo", wsFields), _
                                    GetValue("viagem.favorecido.documento1.numero", wsFields), _
                                    GetValue("viagem.favorecido.documento2.tipo", wsFields), _
                                    GetValue("viagem.favorecido.documento2.numero", wsFields), _
                                    GetValue("viagem.veiculo.qtde", wsFields), _
                                    GetValue("viagem.veiculo1.placa", wsFields), _
                                    GetValue("viagem.veiculo2.placa", wsFields))
            Else
                strXML &= String.Format(strFrota, _
                                    GetValue("viagem.contratante.documento.numero", wsFields), _
                                    GetValue("viagem.favorecido.documento.qtde", wsFields), _
                                    GetValue("viagem.favorecido.documento1.tipo", wsFields), _
                                    GetValue("viagem.favorecido.documento1.numero", wsFields), _
                                    GetValue("viagem.favorecido.documento2.tipo", wsFields), _
                                    GetValue("viagem.favorecido.documento2.numero", wsFields), _
                                    GetValue("viagem.veiculo.qtde", wsFields), _
                                    GetValue("viagem.veiculo1.placa", wsFields))
            End If

            strXML &= "</frota>"

            Dim strNomeArquivo As String = HttpContext.Current.Server.MapPath(String.Format("Pamcard/Envio/findFrota_{0}_{1}.xml", DateTime.Now().ToString("yyyy-MM-dd-HH-mm-ss"), Transportador))

            Dim xmlArquivo As New XmlDocument()
            xmlArquivo.LoadXml(strXML)
            xmlArquivo.Save(strNomeArquivo)
            xmlArquivo = Nothing

            Dim wsResponse As New WSPamcard.ResponseTO()
            wsResponse = wsPamCard.execute(wsRequest)

            Dim wsFieldsResponse() As WSPamcard.FieldTO = wsResponse.fields

            strXML = "<frota>"

            strFrota = "<erro>{0}</erro><errodescricao>{1}</errodescricao><nome>{2}</nome><situacao>{3}</situacao><qtdeplaca>{4}</qtdeplaca><placa1>{5}</placa1><situacaoplaca1>{6}</situacaoplaca1>"

            If Placa.Placa04.Length > 0 Then
                strFrota &= "<placa2>{7}</placa2>"
                strFrota &= "<situacaoplaca2>{8}</situacaoplaca2>"
                strFrota &= "<placa3>{9}</placa3>"
                strFrota &= "<situacaoplaca3>{10}</situacaoplaca3>"
                strFrota &= "<placa4>{11}</placa4>"
                strFrota &= "<situacaoplaca4>{12}</situacaoplaca4>"
                strXML &= String.Format(strFrota, _
                                        GetValue("mensagem.codigo", wsFieldsResponse), _
                                        GetValue("mensagem.descricao", wsFieldsResponse), _
                                        GetValue("viagem.antt.nome", wsFieldsResponse).ToString.Replace("&", "E"), _
                                        GetValue("viagem.antt.rntrc.situacao", wsFieldsResponse), _
                                        GetValue("viagem.veiculo.placa.qdte", wsFieldsResponse), _
                                        GetValue("viagem.veiculo1.placa", wsFieldsResponse), _
                                        GetValue("viagem.veiculo1.situacao", wsFieldsResponse), _
                                        GetValue("viagem.veiculo2.placa", wsFieldsResponse), _
                                        GetValue("viagem.veiculo2.situacao", wsFieldsResponse), _
                                        GetValue("viagem.veiculo3.placa", wsFieldsResponse), _
                                        GetValue("viagem.veiculo3.situacao", wsFieldsResponse), _
                                        GetValue("viagem.veiculo4.placa", wsFieldsResponse), _
                                        GetValue("viagem.veiculo4.situacao", wsFieldsResponse))
            ElseIf Placa.Placa03.Length > 0 Then
                strFrota &= "<placa2>{7}</placa2>"
                strFrota &= "<situacaoplaca2>{8}</situacaoplaca2>"
                strFrota &= "<placa3>{9}</placa3>"
                strFrota &= "<situacaoplaca3>{10}</situacaoplaca3>"
                strXML &= String.Format(strFrota, _
                                        GetValue("mensagem.codigo", wsFieldsResponse), _
                                        GetValue("mensagem.descricao", wsFieldsResponse), _
                                        GetValue("viagem.antt.nome", wsFieldsResponse).ToString.Replace("&", "E"), _
                                        GetValue("viagem.antt.rntrc.situacao", wsFieldsResponse), _
                                        GetValue("viagem.veiculo.placa.qdte", wsFieldsResponse), _
                                        GetValue("viagem.veiculo1.placa", wsFieldsResponse), _
                                        GetValue("viagem.veiculo1.situacao", wsFieldsResponse), _
                                        GetValue("viagem.veiculo2.placa", wsFieldsResponse), _
                                        GetValue("viagem.veiculo2.situacao", wsFieldsResponse), _
                                        GetValue("viagem.veiculo3.placa", wsFieldsResponse), _
                                        GetValue("viagem.veiculo3.situacao", wsFieldsResponse))
            ElseIf Placa.Placa02.Length > 0 Then
                strFrota &= "<placa2>{7}</placa2>"
                strFrota &= "<situacaoplaca2>{8}</situacaoplaca2>"
                strXML &= String.Format(strFrota, _
                                        GetValue("mensagem.codigo", wsFieldsResponse), _
                                        GetValue("mensagem.descricao", wsFieldsResponse), _
                                        GetValue("viagem.antt.nome", wsFieldsResponse).ToString.Replace("&", "E"), _
                                        GetValue("viagem.antt.rntrc.situacao", wsFieldsResponse), _
                                        GetValue("viagem.veiculo.placa.qdte", wsFieldsResponse), _
                                        GetValue("viagem.veiculo1.placa", wsFieldsResponse), _
                                        GetValue("viagem.veiculo1.situacao", wsFieldsResponse), _
                                        GetValue("viagem.veiculo2.placa", wsFieldsResponse), _
                                        GetValue("viagem.veiculo2.situacao", wsFieldsResponse))
            Else
                strXML &= String.Format(strFrota, _
                                        GetValue("mensagem.codigo", wsFieldsResponse), _
                                        GetValue("mensagem.descricao", wsFieldsResponse), _
                                        GetValue("viagem.antt.nome", wsFieldsResponse).ToString.Replace("&", "E"), _
                                        GetValue("viagem.antt.rntrc.situacao", wsFieldsResponse), _
                                        GetValue("viagem.veiculo.placa.qdte", wsFieldsResponse), _
                                        GetValue("viagem.veiculo1.placa", wsFieldsResponse), _
                                        GetValue("viagem.veiculo1.situacao", wsFieldsResponse))
            End If

            strXML &= "</frota>"

            strNomeArquivo = HttpContext.Current.Server.MapPath(String.Format("Pamcard/Retorno/respFrota_{0}_{1}.xml", DateTime.Now().ToString("yyyy-MM-dd-HH-mm-ss"), Transportador))

            xmlArquivo = New XmlDocument()
            xmlArquivo.LoadXml(strXML)
            xmlArquivo.Save(strNomeArquivo)
            xmlArquivo = Nothing

            ds.ReadXml(strNomeArquivo)

            Return ds
        Catch ex As Exception
            HttpContext.Current.Session("ssMessage") = Funcoes.EliminarCaracteresEspeciais(ex.Message.ToString)
            Return ds
        End Try
    End Function

    'Public Function AlterarContratoDeFrete() As String
    '    Dim wsFields(61) As WSPamcard.FieldTO

    '    SetValues("viagem.id", 1026224, wsFields, 0)
    '    'SetValues("viagem.id.cliente", "", wsFields, 0)
    '    SetValues("viagem.antt.ciot.numero", "", wsFields, 1)

    '    SetValues("viagem.contratante.documento.numero", "00001001000312", wsFields, 2)
    '    SetValues("viagem.unidade.documento.tipo", 1, wsFields, 3)
    '    SetValues("viagem.unidade.documento.numero", "00001001000312", wsFields, 4)

    '    'SEM FAVORECIDO
    '    SetValues("viagem.veiculo.qtde", 1, wsFields, 5)
    '    SetValues("viagem.veiculo1.placa", "APX0808", wsFields, 6)
    '    SetValues("viagem.veiculo1.rntrc", "00168810", wsFields, 7)
    '    SetValues("viagem.contrato.numero", "061011", wsFields, 8)

    '    SetValues("viagem.data.partida", "12/10/2011", wsFields, 9)
    '    SetValues("viagem.data.termino", "10/11/2011", wsFields, 10)

    '    SetValues("viagem.carga.natureza", 1212, wsFields, 11)
    '    SetValues("viagem.carga.peso", 1234.55, wsFields, 12)
    '    SetValues("viagem.veiculo.categoria", 7, wsFields, 13)
    '    SetValues("viagem.pedagio.roteirizar", "N", wsFields, 14)
    '    SetValues("viagem.pedagio.solucao.id", 5, wsFields, 15)
    '    SetValues("viagem.pedagio.valor", 10.0, wsFields, 16)
    '    SetValues("viagem.pedagio.cartao", "4417810025749012", wsFields, 17)

    '    SetValues("viagem.documento.qtde", 1, wsFields, 18)
    '    SetValues("viagem.documento1.tipo", 6, wsFields, 19)
    '    SetValues("viagem.documento1.numero", "6554", wsFields, 20)
    '    SetValues("viagem.documento1.serie", 45, wsFields, 21)
    '    SetValues("viagem.documento1.natureza", "1212", wsFields, 22)
    '    SetValues("viagem.documento1.quantidade", "112", wsFields, 23)
    '    SetValues("viagem.documento1.especie", "LIQUIDO", wsFields, 24)
    '    SetValues("viagem.documento1.peso", 1212.65, wsFields, 25)
    '    SetValues("viagem.documento1.cubagem", 9894.55, wsFields, 26)
    '    SetValues("viagem.documento1.mercadoria.valor", 5456.54, wsFields, 27)
    '    SetValues("viagem.documento1.pessoafiscal.qtde", 3, wsFields, 28)
    '    SetValues("viagem.documento1.pessoafiscal1.tipo", 1, wsFields, 29)
    '    SetValues("viagem.documento1.pessoafiscal1.codigo", 2, wsFields, 30)
    '    SetValues("viagem.documento1.pessoafiscal1.documento.tipo", 1, wsFields, 31)
    '    SetValues("viagem.documento1.pessoafiscal1.documento.numero", "27064195000190", wsFields, 32)
    '    SetValues("viagem.documento1.pessoafiscal1.nome", "JENNIFER TRANSPORTES", wsFields, 33)
    '    SetValues("viagem.documento1.pessoafiscal1.endereco.logradouro", "RUA ABILIO SOARES", wsFields, 34)
    '    SetValues("viagem.documento1.pessoafiscal1.endereco.numero", "409", wsFields, 35)
    '    SetValues("viagem.documento1.pessoafiscal1.endereco.complemento", "6 ANDAR", wsFields, 36)
    '    SetValues("viagem.documento1.pessoafiscal1.endereco.bairro", "PARAISO", wsFields, 37)
    '    SetValues("viagem.documento1.pessoafiscal1.endereco.cidade.ibge", "3550308", wsFields, 38)
    '    SetValues("viagem.documento1.pessoafiscal1.endereco.cep", "04005001", wsFields, 39)
    '    SetValues("viagem.documento1.pessoafiscal2.tipo", 2, wsFields, 40)
    '    SetValues("viagem.documento1.pessoafiscal2.codigo", "", wsFields, 41)
    '    SetValues("viagem.documento1.pessoafiscal2.documento.tipo", 2, wsFields, 42)
    '    SetValues("viagem.documento1.pessoafiscal2.documento.numero", "07590701807", wsFields, 43)
    '    SetValues("viagem.documento1.pessoafiscal2.nome", "JENNIFER BARBOSA", wsFields, 44)
    '    SetValues("viagem.documento1.pessoafiscal2.endereco.logradouro", "RUA AUGUSTO GOULART", wsFields, 45)
    '    SetValues("viagem.documento1.pessoafiscal2.endereco.numero", "314", wsFields, 46)
    '    SetValues("viagem.documento1.pessoafiscal2.endereco.complemento", "CASA", wsFields, 47)
    '    SetValues("viagem.documento1.pessoafiscal2.endereco.bairro", "VILA COSTA MELO", wsFields, 48)
    '    SetValues("viagem.documento1.pessoafiscal2.endereco.cidade.ibge", "3550308", wsFields, 49)
    '    SetValues("viagem.documento1.pessoafiscal2.endereco.cep", "03625070", wsFields, 50)
    '    SetValues("viagem.documento1.pessoafiscal3.tipo", 3, wsFields, 51)
    '    SetValues("viagem.documento1.pessoafiscal3.codigo", 2, wsFields, 52)
    '    SetValues("viagem.documento1.pessoafiscal3.documento.tipo", "", wsFields, 53)
    '    SetValues("viagem.documento1.pessoafiscal3.documento.numero", "", wsFields, 54)
    '    SetValues("viagem.documento1.pessoafiscal3.nome", "", wsFields, 55)
    '    SetValues("viagem.documento1.pessoafiscal3.endereco.logradouro", "", wsFields, 56)
    '    SetValues("viagem.documento1.pessoafiscal3.endereco.numero", "", wsFields, 57)
    '    SetValues("viagem.documento1.pessoafiscal3.endereco.complemento", "", wsFields, 58)
    '    SetValues("viagem.documento1.pessoafiscal3.endereco.bairro", "", wsFields, 59)
    '    SetValues("viagem.documento1.pessoafiscal3.endereco.cidade.ibge", "", wsFields, 60)
    '    SetValues("viagem.documento1.pessoafiscal3.endereco.cep", "", wsFields, 61)




    '    Dim wsRequest As New WSPamcard.RequestTO()
    '    '--------------------------------------------------------------------------------------------

    '    Return ""

    'End Function

    'Public Function AlterarValoresDoContratoDeFrete() As String
    '    Dim wsFields(25) As WSPamcard.FieldTO

    '    SetValues("viagem.id", 1026224, wsFields, 0)
    '    'SetValues("viagem.id.cliente", "", wsFields, 0)
    '    SetValues("viagem.antt.ciot.numero", "", wsFields, 1)

    '    SetValues("viagem.contratante.documento.numero", "00001001000312", wsFields, 2)
    '    SetValues("viagem.unidade.documento.tipo", 1, wsFields, 3)
    '    SetValues("viagem.unidade.documento.numero", "00001001000312", wsFields, 4)

    '    SetValues("viagem.parcela.qtde", 2, wsFields, 5)
    '    SetValues("viagem.parcela1.efetivacao.tipo", 1, wsFields, 6)
    '    SetValues("viagem.parcela1.valor", "1.00", wsFields, 7)
    '    SetValues("viagem.parcela1.subtipo", 1, wsFields, 8)
    '    SetValues("viagem.parcela1.base", "N", wsFields, 9)
    '    SetValues("viagem.parcela1.status.id", 1, wsFields, 10)
    '    SetValues("viagem.parcela1.data", "30/10/2011", wsFields, 11)
    '    SetValues("viagem.parcela1.favorecido.tipo.id", 2, wsFields, 12)
    '    SetValues("viagem.parcela1.numero.cliente", 1, wsFields, 13)
    '    SetValues("viagem.parcela2.efetivacao.tipo", 2, wsFields, 14)
    '    SetValues("viagem.parcela2.valor", 2.0, wsFields, 15)
    '    SetValues("viagem.parcela2.subtipo", 2, wsFields, 16)
    '    SetValues("viagem.parcela2.base", "N", wsFields, 17)
    '    SetValues("viagem.parcela2.status.id", 1, wsFields, 18)
    '    SetValues("viagem.parcela2.data", "31/10/2011", wsFields, 19)
    '    SetValues("viagem.parcela2.favorecido.tipo.id", 1, wsFields, 20)
    '    SetValues("viagem.parcela2.numero.cliente", 2, wsFields, 21)

    '    SetValues("viagem.frete.valor.bruto", 200.0, wsFields, 22)
    '    SetValues("viagem.frete.item.qtde", 1, wsFields, 23)
    '    SetValues("viagem.frete.item1.tipo", 1, wsFields, 24)
    '    SetValues("viagem.frete.item1.valor", 189.0, wsFields, 25)



    '    Dim wsRequest As New WSPamcard.RequestTO()
    '    '--------------------------------------------------------------------------------------------

    '    Return ""

    'End Function

    'Private Function UpdateInsertTrip() As String
    '    'Dim wsfields As New WSPamcard.FieldTO(x - 1)

    '    'SetValues("viagem.id.cliente", VIAGEM, wsfields, PosArray)
    '    ''Numero do contrato
    '    'SetValues("viagem.contratante.documento.numero", NRCNPJ, wsfields, PosArray)
    '    ''CNPJ Contratante
    '    'SetValues("viagem.unidade.documento.tipo", 1, wsfields, PosArray)
    '    ''Passado pela HST (CNPJ)
    '    'SetValues("viagem.documento1.tipo", CD_CODIGO_PAMCARD, wsfields, PosArray)
    '    'SetValues("viagem.documento1.numero", NR_SERIE_DOCUMENTACAO, wsfields, PosArray)
    '    'SetValues("viagem.contrato.numero", VIAGEM_PAMCARY, wsfields, PosArray)
    '    ''CD_VIAGEM (COMPLETO)
    '    'Cartao = FindCard(NRCNPJ, NRCARTAO)

    '    ''Popular os dados do cartão
    '    'SetValues("viagem.cartao.tipo", TipoCartao, wsfields, PosArray)
    '    'SetValues("viagem.cartao.numero", NRCARTAO, wsfields, PosArray)

    '    'SetValues("viagem.cartao.portador.documento.tipo", Cartao(2), wsfields, PosArray)
    '    'SetValues("viagem.cartao.portador.documento.numero", Cartao(3), wsfields, PosArray)
    '    'SetValues("viagem.cartao.portador.nome", Cartao(4), wsfields, PosArray)

    '    ''Dados do Veiculo da Viagem
    '    'SetValues("viagem.veiculo.placa", PLACA, wsfields, PosArray)
    '    'SetValues("viagem.veiculo.categoria", "1", wsfields, PosArray)

    '    ''Datas de Partida e chegada
    '    'SetValues("viagem.data.partida", DateTime.Now.ToString("dd/MM/yyyy"), wsfields, PosArray)

    '    ''Dados do Pegadio
    '    'SetValues("viagem.pedagio.origem", "3", wsfields, PosArray)

    '    ''Dados das Parcelas (Valores do contrato)
    '    'SetValues("viagem.parcela.qtde", QtParcelas, wsfields, PosArray)

    '    ''Parcela 1 - Combustivel
    '    'SetValues("viagem.parcela1.numero", "1", wsfields, PosArray)
    '    'SetValues("viagem.parcela1.valor", Valor, wsfields, PosArray)
    '    'SetValues("viagem.parcela1.tipo", "5", wsfields, PosArray)
    '    ''Combustivel
    '    'SetValues("viagem.parcela1.efetivacao.tipo", "1", wsfields, PosArray)
    '    ''Manual
    '    'SetValues("viagem.parcela1.data", DateTime.Now.AddDays(1).ToString("dd/MM/yyyy"), wsfields, PosArray)
    '    ''Dinheiro
    '    'SetValues("viagem.parcela1.status", "2", wsfields, PosArray)
    '    'SetValues("viagem.parcela1.origem", "3", wsfields, PosArray)
    '    'SetValues("viagem.parcela1.base", "S", wsfields, PosArray)

    '    ''Parcela 2 - Dinheiro
    '    'SetValues("viagem.parcela2.numero", "2", wsfields, PosArray)
    '    'SetValues("viagem.parcela2.valor", Saque, wsfields, PosArray)
    '    'SetValues("viagem.parcela2.tipo", "1", wsfields, PosArray)
    '    ''Crédito no Cartão
    '    'SetValues("viagem.parcela2.efetivacao.tipo", "2", wsfields, PosArray)
    '    ''Automatica
    '    'SetValues("viagem.parcela2.data", DateTime.Now.AddDays(1).ToString("dd/MM/yyyy"), wsfields, PosArray)
    '    ''Dinheiro
    '    'SetValues("viagem.parcela2.status", "2", wsfields, PosArray)
    '    'SetValues("viagem.parcela2.origem", "3", wsfields, PosArray)
    '    'SetValues("viagem.parcela2.base", "S", wsfields, PosArray)

    '    ''Parametros para execução do metodo no WS da Pamcary

    '    ''Gerar o Token para Execução
    '    'wsrequest.token = Autentication(CNPJ)

    '    'wsrequest.context = "InsertTrip"
    '    '' ou "UpdateTrip"
    '    'wsrequest.fields = wsfields

    '    ''Executar o metodo no webService
    '    'wsresponse = WSPamcard.execute(wsrequest)

    '    ''Se OK devolve wsresponse.fields[2].key == "viagem.id" e o wsresponse.fields[2].value com o número gerada pela Pamcary
    '    ''Se Não OK devolve (wsresponse.fields[2].key != "viagem.id") o wsresponse.fields[1].value volta com a mensagem de erro

    'End Function

    'Public Function Autentication(ByVal CNPJ As String) As String
    '    'Dim PosArray As Integer = 0

    '    'Dim wsfields As WSPamcard.FieldTO()
    '    'wsfields = New WSPamcard.FieldTO(0) {}

    '    'SetValues("viagem.contratante.documento.numero", CNPJ, wsfields, PosArray)

    '    'wsrequest.certificate = Certificado
    '    ''array de bytes com o certificado do arquivo .CER
    '    'wsrequest.context = "Authentication"
    '    'wsrequest.fields = wsfields

    '    ''Executar o metodo no webService
    '    'wsresponse = wspamcard.execute(wsrequest)

    '    ''Se wsresponse.token != null , devolve o wsresponse.token
    '    ''Se não, pegar msg de erro em wsresponse.fields.GetValue(1)

    'End Function

    'Public Function FindCard(ByVal CNPJ As String, ByVal nrcartao As String) As String()

    '    ''Array dos campos de entrada do webservice  
    '    'Dim wsfields As wsPamcard.FieldTO()
    '    'wsfields = New wsPamcard.FieldTO(3) {}

    '    'SetValues("viagem.contratante.documento.numero", CNPJ, wsfields, PosArray)
    '    'SetValues("viagem.unidade.documento.tipo", 1, wsfields, PosArray)
    '    'SetValues("viagem.cartao.tipo", TipoCartao, wsfields, PosArray)
    '    'SetValues("viagem.cartao.numero", nrcartao, wsfields, PosArray)

    '    ''Gerar o Token para Execução
    '    'wsrequest.token = Autentication(CNPJ)

    '    'wsrequest.certificate = Certificado
    '    'wsrequest.context = "FindCard"
    '    'wsrequest.fields = wsfields

    '    ''Executar o metodo no webService
    '    'wsresponse = wspamcard.execute(wsrequest)

    '    ''Se wsresponse.fields.GetLength(0) < 2, há erro de execução no web service
    '    ''Se wsresponse.fields.GetLength(0) > 2
    '    ''dependendo do tipo do cartão, a posição da msg pode ser estar em wsresponse.fields.GetValue(6) ou wsresponse.fields.GetValue(3)
    '    ''Se a msg for != de "1", erro
    '    ''Se não, cartão OK e os dados do cartão estão 

    'End Function

End Class